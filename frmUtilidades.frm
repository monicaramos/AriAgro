VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUtilidades 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   8040
   Icon            =   "frmUtilidades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8730
      Top             =   5580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCargaCalibres 
      Height          =   4350
      Left            =   0
      TabIndex        =   12
      Top             =   -60
      Width           =   8070
      Begin VB.CheckBox Check2 
         Caption         =   "Actualizar totales"
         Height          =   375
         Left            =   510
         TabIndex        =   19
         Top             =   3210
         Width           =   3525
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   6570
         TabIndex        =   14
         Top             =   3645
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepCarga 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5430
         TabIndex        =   13
         Top             =   3645
         Width           =   975
      End
      Begin VB.Label Label2 
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   600
         TabIndex        =   18
         Top             =   2880
         Width           =   6720
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "NOTA: La tabla facturas_calibre debe de estar vacia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   315
         Index           =   1
         Left            =   570
         TabIndex        =   17
         Top             =   2220
         Width           =   6720
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   $"frmUtilidades.frx":000C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   825
         Index           =   37
         Left            =   480
         TabIndex        =   16
         Top             =   1380
         Width           =   6870
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Carga de Facturas Calibres"
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
         Index           =   0
         Left            =   495
         TabIndex        =   15
         Top             =   495
         Width           =   6735
      End
   End
   Begin VB.Frame FrameInfArticulos 
      Height          =   4350
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8070
      Begin VB.Frame FrameStockMaxMin 
         Caption         =   "Tipo"
         ForeColor       =   &H00972E0B&
         Height          =   1050
         Left            =   480
         TabIndex        =   9
         Top             =   2880
         Width           =   2085
         Begin VB.OptionButton Opcion 
            Caption         =   "Actualizar"
            Height          =   255
            Index           =   1
            Left            =   450
            TabIndex        =   11
            Top             =   570
            Width           =   1245
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Copiar"
            Height          =   255
            Index           =   0
            Left            =   450
            TabIndex        =   10
            Top             =   270
            Width           =   1305
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Calibres"
         Height          =   195
         Index           =   0
         Left            =   540
         TabIndex        =   8
         Top             =   2160
         Width           =   2130
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Calidades"
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   7
         Top             =   2520
         Width           =   2130
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   70
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Text5"
         Top             =   1620
         Width           =   4305
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   70
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1620
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5400
         TabIndex        =   2
         Top             =   3645
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   6570
         TabIndex        =   3
         Top             =   3645
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Copia de Calibres / Calidades a Variedad"
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
         TabIndex        =   6
         Top             =   495
         Width           =   6735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   27
         Left            =   1425
         Top             =   1620
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Variedad Destino"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   38
         Left            =   510
         TabIndex        =   5
         Top             =   1350
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmUtilidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NumCod As String 'Variedad origen

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto
Public Opcionlistado As Long
' 0- Carga de los registros de facturas_calibre
' 1- Modificacion del codigo de socio en alzira segun una tabla (socio 8003)

Private HaDevueltoDatos As Boolean

Private WithEvents frmVar As frmManVariedad
Attribute frmVar.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim indFrame As Single
Dim indCodigo As Integer

Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub CmdAcepCarga_Click()

    If CargarFacturasCalibres Then
        MsgBox "Proceso realizado correctamente.", vbExclamation
    End If

End Sub

Private Function CargarFacturasCalibres()
Dim sql As String
Dim Sql2 As String
Dim Mens As String
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim cadWHERE As String

    On Error GoTo eCargarFacturasCalibres
    
    sql = "select * from facturas_variedad order by codtipom , numfactu, fecfactu"
    
    conn.BeginTrans
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    b = True
    
    
    While Not Rs.EOF And b
        Mens = "Insertar Calibres"
        
        Label2(2).Caption = "Procesando Factura: " & DBLet(Rs!codTipoM, "T") & "-" & CStr(DBLet(Rs!NumFactu, "N")) & " de " & CStr(DBLet(Rs!FecFactu, "F"))
        DoEvents
        
        b = InsertarModificarCalibres(True, DBLet(Rs!codTipoM, "T"), CStr(DBLet(Rs!NumFactu, "N")), CStr(DBLet(Rs!FecFactu, "F")), CStr(DBLet(Rs!numlinea, "N")), CStr(DBLet(Rs!NumAlbar, "N")), CStr(DBLet(Rs!numlinealbar, "N")), CStr(DBLet(Rs!cantreal, "N")), CStr(DBLet(Rs!Unidades, "N")), CStr(DBLet(Rs!imporbru, "N")), CStr(DBLet(Rs!impornet, "N")), Mens, CStr(DBLet(Rs!cantfact, "N")))
        
        If Check2.Value Then
            cadWHERE = "facturas.codtipom = " & DBSet(Rs!codTipoM, "T") & " and facturas.numfactu = " & DBSet(Rs!NumFactu, "N") & " and facturas.fecfactu = " & DBSet(Rs!FecFactu, "F")
            If b Then b = CalcularDatosFacturaVenta(cadWHERE)
        End If
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing

eCargarFacturasCalibres:
    If Err.Number <> 0 Then
        Mens = Mens & vbCrLf & Err.Description
        b = False
    End If
    If b Then
        CargarFacturasCalibres = True
        conn.CommitTrans
    Else
        CargarFacturasCalibres = False
        conn.RollbackTrans
    End If
End Function

Private Function InsertarModificarCalibres(Insertar As Boolean, codTipoM As String, Factura As String, FecFactu As String, numlinea As String, Albaran As String, NumlineaAlb As String, TCantReal As String, TUnidades As String, TImpBruto As String, TImpNeto As String, MenError As String, TCantFact As String) As Boolean
' Insertar : = true : insertamos todas las lineas en facturas_calibre del albaran prorrateando
'            = false: venimos de modificar lineas en facturas_variedad prorrateamos lineas de facturas_calibre segun los cambios que hay en facturas_variedad
Dim Rs As ADODB.Recordset
Dim sql As String
Dim Sql2 As String
Dim vImpDto As Currency
Dim vDto1 As Currency
Dim vDto2 As Currency
Dim vImpNeto As Currency
Dim vImpBruto As Currency
Dim vPrecNeto As Currency
Dim vPrecBruto As Currency

Dim TipoDto As String
Dim ImpDto As String
Dim Cliente As String
Dim Rdo As Long

Dim ImpBrutoAc As Currency
Dim ImpNetoAc As Currency

Dim Diferencia As Currency
Dim Diferencia1 As Currency

Dim UltimaLinea As Currency
Dim TipoFactFor As Byte

Dim vHayReg As Byte
Dim KilosCaja As Currency

    On Error GoTo eInsertarModificarCalibres


    KilosCaja = DevuelveValor("select kiloscaj from forfaits inner join albaran_variedad on forfaits.codforfait = albaran_variedad.codforfait where albaran_variedad.numalbar = " & DBSet(Albaran, "N") & " and numlinea = " & DBSet(NumlineaAlb, "N"))

    ' Si venimos de insertar una linea de factura, insertamos automaticamente todas las lineas de calibre prorrateando
    If Insertar Then
        ' Primero insertamos con los precios e importes a cero
        
'        sql = "insert into facturas_calibre (codtipom,numfactu,fecfactu,numlinea,numline1,numalbar,numlinealbar,numline1albar,cantreal,cantfact,"
'        sql = sql & " precibru,precinet,dtocom1,dtocom2,imporbru,impornet,unidades) "
'        sql = sql & " select " & DBSet(CodTipoM, "T") & "," & DBSet(Factura, "N") & "," & DBSet(FecFactu, "F") & ","
'        sql = sql & DBSet(numlinea, "N") & ",numline1," & DBSet(Albaran, "N") & "," & DBSet(NumlineaAlb, "N") & ",numline1,"
'        sql = sql & " pesoneto, round(numcajas * " & DBSet(KilosCaja, "N") & ",2), 0,0,0,0,0,0,unidades "
'        sql = sql & " from albaran_calibre where numalbar = " & DBSet(Albaran, "N")
'        sql = sql & " and numlinea = " & DBSet(NumlineaAlb, "N")
'        sql = sql & " order by numline1 "
'
'        conn.Execute sql
        
        'prorrateamos las cantidades con respecto al peso neto
        Dim TPesoNeto As String
        Dim vCantReal As String
        Dim vCantFact As String
        Dim CantRealAc As Currency
        Dim CantFactAc As Currency
        Dim Linea1 As Long
        
        sql = " select " & DBSet(codTipoM, "T") & "," & DBSet(Factura, "N") & "," & DBSet(FecFactu, "F") & ","
        sql = sql & DBSet(numlinea, "N") & ",numline1," & DBSet(Albaran, "N") & "," & DBSet(NumlineaAlb, "N") & ",numline1,"
        sql = sql & " pesoneto, numcajas, 0,0,0,0,0,0,unidades "
        sql = sql & " from albaran_calibre where numalbar = " & DBSet(Albaran, "N")
        sql = sql & " and numlinea = " & DBSet(NumlineaAlb, "N")
        sql = sql & " order by numline1 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            TPesoNeto = DevuelveValor("select sum(pesoneto) from albaran_calibre where numalbar = " & DBSet(Albaran, "N") & " and numlinea = " & DBSet(NumlineaAlb, "N"))
            vCantReal = 0
            vCantFact = 0
            If TPesoNeto <> 0 Then
                vCantReal = Round2(TCantReal * DBLet(Rs!Pesoneto, "N") / TPesoNeto, 0)
                vCantFact = Round2(TCantFact * DBLet(Rs!Pesoneto, "N") / TPesoNeto, 0)
                
                CantRealAc = CantRealAc + vCantReal
                CantFactAc = CantFactAc + vCantFact
            End If
        
            sql = "insert into facturas_calibre (codtipom,numfactu,fecfactu,numlinea,numline1,numalbar,numlinealbar,numline1albar,cantreal,cantfact,"
            sql = sql & " precibru,precinet,dtocom1,dtocom2,imporbru,impornet,unidades)  values ("
            sql = sql & DBSet(codTipoM, "T") & "," & DBSet(Factura, "N") & "," & DBSet(FecFactu, "F") & ","
            sql = sql & DBSet(numlinea, "N") & "," & DBSet(Rs!numline1, "N") & "," & DBSet(Albaran, "N") & "," & DBSet(NumlineaAlb, "N") & "," & DBSet(Rs!numline1, "N") & ","
            sql = sql & DBSet(vCantReal, "N") & "," & DBSet(vCantFact, "N") & ",0,0,0,0,0,0," & DBSet(Rs!Unidades, "N") & ")"
    
            Linea1 = DBLet(Rs!numline1, "N")
            
            conn.Execute sql
        
        
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
        ' en caso de que hayan descuadres
        Dim DiferenciaReal As Currency
        Dim DiferenciaFact As Currency
        
        DiferenciaReal = TCantReal - CantRealAc
        DiferenciaFact = TCantFact - CantFactAc
        If CantRealAc <> TCantReal Or CantFactAc <> TCantFact Then
            sql = "update facturas_calibre set cantreal = cantreal + " & DBSet(DiferenciaReal, "N") & ","
            sql = sql & " cantfact = cantfact + " & DBSet(DiferenciaFact, "N")
            sql = sql & " where codtipom = " & DBSet(codTipoM, "T") & " and numfactu = " & DBSet(Factura, "N")
            sql = sql & " and numlinea = " & DBSet(numlinea, "N")
            sql = sql & " and numline1 = " & DBSet(Linea1, "N")
        
            conn.Execute sql
        End If

        
    End If
    
    ' Prorrateamos TODO con respecto a los kilos
    sql = "select * from facturas_calibre where codtipom = " & DBSet(codTipoM, "T") & " and numfactu = " & DBSet(Factura, "N")
    sql = sql & " and fecfactu = " & DBSet(FecFactu, "F") & " and numlinea = " & DBSet(numlinea, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    sql = ""
    sql = DevuelveDesdeBDNew(cAgro, "facturas", "impdtoc", "codtipom", codTipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vImpDto = ComprobarCero(sql)
    
    sql = ""
    sql = DevuelveDesdeBDNew(cAgro, "facturas", "dtocom1", "codtipom", codTipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vDto1 = ComprobarCero(sql)
    
    sql = ""
    sql = DevuelveDesdeBDNew(cAgro, "facturas", "dtocom2", "codtipom", codTipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vDto2 = ComprobarCero(sql)
    
    '++monica:030608:traemos el redondeo del precio
    sql = ""
    sql = DevuelveDesdeBDNew(cAgro, "facturas", "codclien", "codtipom", codTipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    Cliente = ComprobarCero(sql)
    sql = ""
    sql = DevuelveDesdeBDNew(cAgro, "clientes", "nrodecprec", "codclien", Cliente, "N")
    Rdo = ComprobarCero(sql)
    
    Dim vPrecio As Currency
    Dim Rdo1 As Integer
    Dim Rdo2 As Integer
    vPrecio = DevuelveValor("select precinet from facturas_variedad where codtipom = " & DBSet(codTipoM, "T") & " and numfactu = " & DBSet(Factura, "N") & " and fecfactu = " & DBSet(FecFactu, "F") & " and numlinea = " & DBSet(numlinea, "N"))
    
    Rdo1 = 4
    If (vPrecio * 10) - Int(vPrecio * 10) <> 0 Then Rdo1 = 2
    If (vPrecio * 100) - Int(vPrecio * 100) <> 0 Then Rdo1 = 3
    If (vPrecio * 1000) - Int(vPrecio * 1000) <> 0 Then Rdo1 = 4
    
    vPrecio = DevuelveValor("select precibru from facturas_variedad where codtipom = " & DBSet(codTipoM, "T") & " and numfactu = " & DBSet(Factura, "N") & " and fecfactu = " & DBSet(FecFactu, "F") & " and numlinea = " & DBSet(numlinea, "N"))
    
    Rdo2 = 4
    If (vPrecio * 10) - Int(vPrecio * 10) <> 0 Then Rdo2 = 2
    If (vPrecio * 100) - Int(vPrecio * 100) <> 0 Then Rdo2 = 3
    If (vPrecio * 1000) - Int(vPrecio * 1000) <> 0 Then Rdo2 = 4
    
    
    If Rdo1 > Rdo2 Then
        Rdo = Rdo1
    Else
        Rdo = Rdo2
    End If
    
    vHayReg = 0
    
    
    ImpBrutoAc = 0
    ImpNetoAc = 0
    
    While Not Rs.EOF
        vHayReg = 1
        
        TipoDto = DevuelveDesdeBDNew(cAgro, "clientes", "tipodtos", "codclien", Cliente, "N")
        If TipoFacturarForfaits(CStr(Albaran), CStr(NumlineaAlb)) = 1 Then 'kilos
            TipoFactFor = 1
            
            vImpBruto = Round2(TImpBruto * DBLet(Rs!cantreal, "N") / TCantReal, 2)
            
            ImpBrutoAc = ImpBrutoAc + vImpBruto
            
            vImpNeto = Round2(TImpNeto * DBLet(Rs!cantreal, "N") / TCantReal, 2)
            
            
            ImpNetoAc = ImpNetoAc + vImpNeto
                
            '[Monica]24/11/2011: si las unidades son 0 no hay division
            'precio neto
            vPrecNeto = 0
            vPrecBruto = 0
            If DBLet(Rs!cantreal, "N") <> 0 Then
                vPrecNeto = Round2(vImpNeto / DBLet(Rs!cantreal, "N"), Rdo)
                vPrecBruto = Round2(vImpBruto / DBLet(Rs!cantreal, "N"), Rdo)
            End If
            '++monica:040608 : solo si redondeo <> 4
'            If Rdo = 2 Or Rdo = 3 Then
'                vImpNeto = Round2(vPrecNeto * DBLet(Rs!cantreal, "N"), 2)
'                vImpBruto = Round2(vPrecBruto * DBLet(Rs!cantreal, "N"), 2)
'            End If
            
        Else 'unidades
            TipoFactFor = 0
'            ImpDto = CalcularImporteDto(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), TipoM, Factura, FecFactu, CStr(vImpDto), True)
'            vImpNeto = CalcularImporteFClien(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), CStr(vDto1), CStr(vDto2), CByte(TipoDto), CStr(ImpDto), DBLet(Rs!imporbru, "N"))

            vImpBruto = 0
            If TUnidades <> 0 Then
                vImpBruto = Round2(TImpBruto * DBLet(Rs!Unidades, "N") / TUnidades, 2)
            End If
            ImpBrutoAc = ImpBrutoAc + vImpBruto
            
            vImpNeto = 0
            If TUnidades <> 0 Then
                vImpNeto = Round2(TImpNeto * DBLet(Rs!Unidades, "N") / TUnidades, 2)
            End If
            ImpNetoAc = ImpNetoAc + vImpNeto
            
            '[Monica]24/11/2011: si las unidades son 0 no hay division
            'precio neto
            vPrecNeto = 0
            vPrecBruto = 0
            If DBLet(Rs!Unidades, "N") <> 0 Then
                vPrecNeto = Round2(vImpNeto / DBLet(Rs!Unidades, "N"), Rdo)
                vPrecBruto = Round2(vImpBruto / DBLet(Rs!Unidades, "N"), Rdo)
            End If
            
            '++monica:040608
            If Rdo = 2 Or Rdo = 3 Then
                vImpNeto = Round2(vPrecNeto * DBLet(Rs!Unidades, "N"), 2)
                vImpBruto = Round2(vPrecBruto * DBLet(Rs!Unidades, "N"), 2)
            End If
            
        End If
        
        
        Sql2 = "update facturas_calibre set "
        Sql2 = Sql2 & "precibru = " & DBSet(vPrecBruto, "N")
        Sql2 = Sql2 & ",precinet = " & DBSet(vPrecNeto, "N")
        Sql2 = Sql2 & ",imporbru = " & DBSet(vImpBruto, "N")
        Sql2 = Sql2 & ",impornet = " & DBSet(vImpNeto, "N")
        Sql2 = Sql2 & ",dtocom1 = " & DBSet(vDto1, "N")
        Sql2 = Sql2 & ",dtocom2 = " & DBSet(vDto2, "N")
        Sql2 = Sql2 & " where codtipom = " & DBSet(codTipoM, "T")
        Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
        Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
        Sql2 = Sql2 & " and numlinea = " & DBSet(numlinea, "N")
        Sql2 = Sql2 & " and numline1 = " & DBSet(Rs!numline1, "N")
    
        conn.Execute Sql2
    
        UltimaLinea = DBLet(Rs!numline1, "N")
    
        Rs.MoveNext
    Wend
    
    Rs.Close
    
    '[Monica]16/09/2011: si no coincide la suma con los totales redondeamos en la ultima linea
    If vHayReg = 1 Then
        If ImpBrutoAc <> TImpBruto Or ImpNetoAc <> TImpNeto Then
            Diferencia = TImpBruto - ImpBrutoAc
            Diferencia1 = TImpNeto - ImpNetoAc
            
            Sql2 = "update facturas_calibre set impornet = impornet + " & DBSet(Diferencia1, "N")
            Sql2 = Sql2 & ", imporbru = imporbru + " & DBSet(Diferencia, "N")
            Sql2 = Sql2 & " where codtipom = " & DBSet(codTipoM, "T")
            Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
            Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
            Sql2 = Sql2 & " and numlinea = " & DBSet(numlinea, "N")
            Sql2 = Sql2 & " and numline1 = " & DBSet(UltimaLinea, "N")
        
            conn.Execute Sql2
    
            If TipoFactFor = 1 Then 'kilos
                Sql2 = "update facturas_calibre set precinet = round(impornet / cantreal, " & DBSet(Rdo, "N") & ") "
                Sql2 = Sql2 & " where codtipom = " & DBSet(codTipoM, "T")
                Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
                Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
                Sql2 = Sql2 & " and numlinea = " & DBSet(numlinea, "N")
                Sql2 = Sql2 & " and numline1 = " & DBSet(UltimaLinea, "N")
            
                conn.Execute Sql2
            Else 'unidades
                'precio neto
                Sql2 = "update facturas_calibre set precinet = round(impornet / unidades, " & DBSet(Rdo, "N") & ") "
                Sql2 = Sql2 & " where codtipom = " & DBSet(codTipoM, "T")
                Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
                Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
                Sql2 = Sql2 & " and numlinea = " & DBSet(numlinea, "N")
                Sql2 = Sql2 & " and numline1 = " & DBSet(UltimaLinea, "N")
            
                conn.Execute Sql2
            End If
        End If
    End If
    
    Set Rs = Nothing
    
    InsertarModificarCalibres = True
    Exit Function

eInsertarModificarCalibres:
    If Err.Number <> 0 Then
        MenError = MenError & vbCrLf & Err.Description
        InsertarModificarCalibres = False
    End If
End Function

'Private Function InsertarModificarCalibres(Insertar As Boolean, CodTipoM As String, Factura As String, FecFactu As String, numlinea As String, Albaran As String, NumlineaAlb As String, TCantReal As String, TUnidades As String, TImpBruto As String, TImpNeto As String, TCantFact As String, MenError As String) As Boolean
'' Insertar : = true : insertamos todas las lineas en facturas_calibre del albaran prorrateando
''            = false: venimos de modificar lineas en facturas_variedad prorrateamos lineas de facturas_calibre segun los cambios que hay en facturas_variedad
'Dim RS As ADODB.Recordset
'Dim sql As String
'Dim Sql2 As String
'Dim vImpDto As Currency
'Dim vDto1 As Currency
'Dim vDto2 As Currency
'Dim vImpNeto As Currency
'Dim vImpBruto As Currency
'Dim vPrecNeto As Currency
'Dim vPrecBruto As Currency
'
'Dim TipoDto As String
'Dim ImpDto As String
'Dim Cliente As String
'Dim Rdo As Long
'
'Dim ImpBrutoAc As Currency
'Dim ImpNetoAc As Currency
'
'Dim Diferencia As Currency
'Dim Diferencia1 As Currency
'Dim Diferencia2 As Currency
'
'Dim UltimaLinea As Currency
'Dim TipoFactFor As Byte
'
'Dim vHayReg As Byte
'Dim KilosCaja As Currency
'Dim CantFactAc As Long
'Dim vCantFact As Long
'Dim TNumCaja As Long
'Dim vNumCaja As Long
'
'    On Error GoTo eInsertarModificarCalibres
'
'
'    KilosCaja = DevuelveValor("select kiloscaj from forfaits inner join albaran_variedad on forfaits.codforfait = albaran_variedad.codforfait where albaran_variedad.numalbar = " & DBSet(Albaran, "N") & " and numlinea = " & DBSet(NumlineaAlb, "N"))
'
'    ' Si venimos de insertar una linea de factura, insertamos automaticamente todas las lineas de calibre prorrateando
'    If Insertar Then
''        ' Primero insertamos con los precios e importes a cero
''
''        sql = "insert into facturas_calibre (codtipom,numfactu,fecfactu,numlinea,numline1,numalbar,numlinealbar,numline1albar,cantreal,cantfact,"
''        sql = sql & " precibru,precinet,dtocom1,dtocom2,imporbru,impornet,unidades) "
''        sql = sql & " select " & DBSet(CodTipoM, "T") & "," & DBSet(Factura, "N") & "," & DBSet(FecFactu, "F") & ","
''        sql = sql & DBSet(numlinea, "N") & ",numline1," & DBSet(Albaran, "N") & "," & DBSet(NumlineaAlb, "N") & ",numline1,"
''        sql = sql & " pesoneto, 0, 0,0,0,0,0,0,unidades "
''        sql = sql & " from albaran_calibre where numalbar = " & DBSet(Albaran, "N")
''        sql = sql & " and numlinea = " & DBSet(NumlineaAlb, "N")
''        sql = sql & " order by numline1 "
''
''        conn.Execute sql
'
'        'prorrateamos las cantidades con respecto al peso neto
'        Dim TPesoneto As String
'        Dim vCantReal As String
'        Dim CantRealAc As Currency
'        Dim Linea1 As Long
'
'
'        sql = " select " & DBSet(CodTipoM, "T") & "," & DBSet(Factura, "N") & "," & DBSet(FecFactu, "F") & ","
'        sql = sql & DBSet(numlinea, "N") & ",numline1," & DBSet(Albaran, "N") & "," & DBSet(NumlineaAlb, "N") & ",numline1,"
'        sql = sql & " pesoneto, numcajas, 0,0,0,0,0,0,unidades "
'        sql = sql & " from albaran_calibre where numalbar = " & DBSet(Albaran, "N")
'        sql = sql & " and numlinea = " & DBSet(NumlineaAlb, "N")
'        sql = sql & " order by numline1 "
'
'        Set RS = New ADODB.Recordset
'        RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        While Not RS.EOF
'            TPesoneto = DevuelveValor("select sum(pesoneto) from albaran_calibre where numalbar = " & DBSet(Albaran, "N") & " and numlinea = " & DBSet(NumlineaAlb, "N"))
'            vCantReal = 0
'            vCantFact = 0
'            If TPesoneto <> 0 Then
'                vCantReal = Round2(TCantReal * DBLet(RS!Pesoneto, "N") / TPesoneto, 0)
'                vCantFact = Round2(TCantReal * DBLet(RS!Pesoneto, "N") / TPesoneto, 0)
'
'                CantRealAc = CantRealAc + vCantReal
'                CantFactAc = CantFactAc + vCantFact
'            End If
'
'            sql = "insert into facturas_calibre (codtipom,numfactu,fecfactu,numlinea,numline1,numalbar,numlinealbar,numline1albar,cantreal,cantfact,"
'            sql = sql & " precibru,precinet,dtocom1,dtocom2,imporbru,impornet,unidades)  values ("
'            sql = sql & DBSet(CodTipoM, "T") & "," & DBSet(Factura, "N") & "," & DBSet(FecFactu, "F") & ","
'            sql = sql & DBSet(numlinea, "N") & "," & DBSet(RS!numline1, "N") & "," & DBSet(Albaran, "N") & "," & DBSet(NumlineaAlb, "N") & "," & DBSet(RS!numline1, "N") & ","
'            sql = sql & DBSet(vCantReal, "N") & "," & DBSet(vCantFact, "N") & ",0,0,0,0,0,0," & DBSet(RS!Unidades, "N") & ")"
'
'            conn.Execute sql
'
'
'            RS.MoveNext
'        Wend
'        Set RS = Nothing
'
'
'        ' en caso de que hayan descuadres
'        Dim DiferenciaReal As Currency
'        Dim DiferenciaFact As Currency
'
'        DiferenciaReal = TCantReal - CantRealAc
'        DiferenciaFact = TCantFact - CantFactAc
'        If CantRealAc <> TCantReal Or CantFactAc <> TCantFact Then
'            sql = "update facturas_calibre set cantreal = cantreal + " & DBSet(DiferenciaReal, "N") & ","
'            sql = sql & " cantfact = cantfact + " & DBSet(DiferenciaFact, "N")
'            sql = sql & " where codtipom = " & DBSet(CodTipoM, "T") & " and numfactu = " & DBSet(Factura, "N")
'            sql = sql & " and numlinea = " & DBSet(numlinea, "N")
'            sql = sql & " and numline1 = " & DBSet(Linea1, "N")
'
'            conn.Execute sql
'        End If
'
'    End If
'
'    ' Prorrateamos TODO con respecto a los kilos
'    sql = "select * from facturas_calibre where codtipom = " & DBSet(CodTipoM, "T") & " and numfactu = " & DBSet(Factura, "N")
'    sql = sql & " and fecfactu = " & DBSet(FecFactu, "F") & " and numlinea = " & DBSet(numlinea, "N")
'
'    Set RS = New ADODB.Recordset
'    RS.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    sql = ""
'    sql = DevuelveDesdeBDNew(cAgro, "facturas", "impdtoc", "codtipom", CodTipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
'    vImpDto = ComprobarCero(sql)
'
'    sql = ""
'    sql = DevuelveDesdeBDNew(cAgro, "facturas", "dtocom1", "codtipom", CodTipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
'    vDto1 = ComprobarCero(sql)
'
'    sql = ""
'    sql = DevuelveDesdeBDNew(cAgro, "facturas", "dtocom2", "codtipom", CodTipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
'    vDto2 = ComprobarCero(sql)
'
'    '++monica:030608:traemos el redondeo del precio
'    sql = ""
'    sql = DevuelveDesdeBDNew(cAgro, "facturas", "codclien", "codtipom", CodTipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
'    Cliente = ComprobarCero(sql)
'    sql = ""
'    sql = DevuelveDesdeBDNew(cAgro, "clientes", "nrodecprec", "codclien", Cliente, "N")
'    Rdo = ComprobarCero(sql)
'
'    vHayReg = 0
'
'    ImpBrutoAc = 0
'    ImpNetoAc = 0
'    CantFactAc = 0
'    While Not RS.EOF
'        vHayReg = 1
'
''        TNumCaja = DevuelveValor("select numcajas from albaran_variedad where numalbar = " & DBSet(Albaran, "N") & " and numlinea = " & DBSet(NumlineaAlb, "N"))
''
''        vCantFact = 0
''        If TNumCaja <> 0 Then
''            vNumCaja = DevuelveValor("select numcajas from albaran_calibre where numalbar = " & DBSet(Albaran, "N") & " and numlinea = " & DBSet(NumlineaAlb, "N") & "  and numline1 = " & DBSet(RS!numline1, "N"))
''
''            vCantFact = Round2(TCantFact * vNumCaja / TNumCaja, 0)
''        End If
''        CantFactAc = CantFactAc + vCantFact
'
'        TipoDto = DevuelveDesdeBDNew(cAgro, "clientes", "tipodtos", "codclien", Cliente, "N")
'        If TipoFacturarForfaits(CStr(Albaran), CStr(NumlineaAlb)) = 1 Then 'kilos
'            TipoFactFor = 1
'
'            vImpBruto = Round2(TImpBruto * DBLet(RS!cantreal, "N") / TCantReal, 2)
'            ImpBrutoAc = ImpBrutoAc + vImpBruto
'
'            vImpNeto = Round2(TImpNeto * DBLet(RS!cantreal, "N") / TCantReal, 2)
'            ImpNetoAc = ImpNetoAc + vImpNeto
'
'            '[Monica]24/11/2011: si las unidades son 0 no hay division
'            'precio neto
'            vPrecNeto = 0
'            vPrecBruto = 0
'            If DBLet(RS!cantreal, "N") <> 0 Then
'                vPrecNeto = Round2(vImpNeto / DBLet(RS!cantreal, "N"), Rdo)
'                vPrecBruto = Round2(vImpBruto / DBLet(RS!cantreal, "N"), Rdo)
'            End If
'            '++monica:040608 : solo si redondeo <> 4
'            If Rdo = 2 Or Rdo = 3 Then
'                vImpNeto = Round2(vPrecNeto * DBLet(RS!cantreal, "N"), 2)
'            End If
'
'        Else 'unidades
'            TipoFactFor = 0
''            ImpDto = CalcularImporteDto(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), TipoM, Factura, FecFactu, CStr(vImpDto), True)
''            vImpNeto = CalcularImporteFClien(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), CStr(vDto1), CStr(vDto2), CByte(TipoDto), CStr(ImpDto), DBLet(Rs!imporbru, "N"))
'
'            vImpBruto = 0
'            If TUnidades <> 0 Then
'                vImpBruto = Round2(TImpBruto * DBLet(RS!Unidades, "N") / TUnidades, 2)
'            End If
'            ImpBrutoAc = ImpBrutoAc + vImpBruto
'
'            vImpNeto = 0
'            If TUnidades <> 0 Then
'                vImpNeto = Round2(TImpNeto * DBLet(RS!Unidades, "N") / TUnidades, 2)
'            End If
'            ImpNetoAc = ImpNetoAc + vImpNeto
'
'            '[Monica]24/11/2011: si las unidades son 0 no hay division
'            'precio neto
'            vPrecNeto = 0
'            vPrecBruto = 0
'            If DBLet(RS!Unidades, "N") <> 0 Then
'                vPrecNeto = Round2(vImpNeto / DBLet(RS!Unidades, "N"), Rdo)
'                vPrecBruto = Round2(vImpBruto / DBLet(RS!Unidades, "N"), Rdo)
'            End If
'
'            '++monica:040608
'            If Rdo = 2 Or Rdo = 3 Then
'                vImpNeto = Round2(vPrecNeto * DBLet(RS!Unidades, "N"), 2)
'                vImpBruto = Round2(vPrecBruto * DBLet(RS!Unidades, "N"), 2)
'            End If
'        End If
'
'        Sql2 = "update facturas_calibre set "
'        Sql2 = Sql2 & "precibru = " & DBSet(vPrecBruto, "N")
'        Sql2 = Sql2 & ",precinet = " & DBSet(vPrecNeto, "N")
'        Sql2 = Sql2 & ",imporbru = " & DBSet(vImpBruto, "N")
'        Sql2 = Sql2 & ",impornet = " & DBSet(vImpNeto, "N")
'        Sql2 = Sql2 & ",dtocom1 = " & DBSet(vDto1, "N")
'        Sql2 = Sql2 & ",dtocom2 = " & DBSet(vDto2, "N")
'        Sql2 = Sql2 & ",cantfact = " & DBSet(vCantFact, "N")
'        Sql2 = Sql2 & " where codtipom = " & DBSet(CodTipoM, "T")
'        Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
'        Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
'        Sql2 = Sql2 & " and numlinea = " & DBSet(numlinea, "N")
'        Sql2 = Sql2 & " and numline1 = " & DBSet(RS!numline1, "N")
'
'        conn.Execute Sql2
'
'        UltimaLinea = DBLet(RS!numline1, "N")
'
'        RS.MoveNext
'    Wend
'
'    RS.Close
'
'    '[Monica]16/09/2011: si no coincide la suma con los totales redondeamos en la ultima linea
'    If vHayReg = 1 Then
'        If ImpBrutoAc <> TImpBruto Or ImpNetoAc <> TImpNeto Or CantFactAc <> TCantFact Then
'            Diferencia = TImpBruto - ImpBrutoAc
'            Diferencia1 = TImpNeto - ImpNetoAc
'            Diferencia2 = TCantFact - CantFactAc
'
'            Sql2 = "update facturas_calibre set impornet = impornet + " & DBSet(Diferencia1, "N")
'            Sql2 = Sql2 & ", imporbru = imporbru + " & DBSet(Diferencia, "N")
'            Sql2 = Sql2 & ", cantfact = cantfact + " & DBSet(Diferencia2, "N")
'            Sql2 = Sql2 & " where codtipom = " & DBSet(CodTipoM, "T")
'            Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
'            Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
'            Sql2 = Sql2 & " and numlinea = " & DBSet(numlinea, "N")
'            Sql2 = Sql2 & " and numline1 = " & DBSet(UltimaLinea, "N")
'
'            conn.Execute Sql2
'
'            If TipoFactFor = 1 Then 'kilos
'                Sql2 = "update facturas_calibre set precinet = round(impornet / cantreal, " & DBSet(Rdo, "N") & ") "
'                Sql2 = Sql2 & " where codtipom = " & DBSet(CodTipoM, "T")
'                Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
'                Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
'                Sql2 = Sql2 & " and numlinea = " & DBSet(numlinea, "N")
'                Sql2 = Sql2 & " and numline1 = " & DBSet(UltimaLinea, "N")
'
'                conn.Execute Sql2
'            Else 'unidades
'                'precio neto
'                Sql2 = "update facturas_calibre set precinet = round(impornet / unidades, " & DBSet(Rdo, "N") & ") "
'                Sql2 = Sql2 & " where codtipom = " & DBSet(CodTipoM, "T")
'                Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
'                Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
'                Sql2 = Sql2 & " and numlinea = " & DBSet(numlinea, "N")
'                Sql2 = Sql2 & " and numline1 = " & DBSet(UltimaLinea, "N")
'
'                conn.Execute Sql2
'            End If
'        End If
'    End If
'
'    Set RS = Nothing
'
'    InsertarModificarCalibres = True
'    Exit Function
'
'eInsertarModificarCalibres:
'    If Err.Number <> 0 Then
'        MenError = MenError & vbCrLf & Err.Description
'        InsertarModificarCalibres = False
'    End If
'End Function


Private Sub cmdAceptar_Click()
Dim sql As String
Dim SQL1 As String
    
    
    
'    If BloqueaRegistro("variedades", "codvarie = " & DBSet(txtCodigo(70).Text, "N")) Then

    If DatosOk Then
        If ActualizarRegistros Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
            cmdCancel_Click
        End If
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(70)
        Me.Opcion(0).Value = True
        Check1(0).Value = 1
        Check1(1).Value = 1
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

    For i = 27 To 27
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'    Me.Width = W + 70
'    Me.Height = H + 350
End Sub



Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de variedades
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)
'Buscar general: cada index llama a una tabla
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 27 'cod. de variedad
            indCodigo = 70
            Set frmVar = New frmManVariedad
            frmVar.DatosADevolverBusqueda = "0|1|" 'Abrimos en Modo Busqueda
            frmVar.DeConsulta = True
            frmVar.Show vbModal
            Set frmVar = Nothing
            
    End Select
    PonerFoco txtCodigo(indCodigo)
    Screen.MousePointer = vbDefault
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
        
    Select Case Index
        Case 70  'Cod.variedad
            If txtCodigo(Index).Text <> "" Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
                If txtNombre(Index).Text = "" Then
                    MsgBox "Variedad no existe. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(Index)
                End If
            End If
    End Select
    
End Sub

Private Function DatosOk() As Boolean
Dim i As Integer

    DatosOk = False
    If txtCodigo(70).Text = "" Then
        MsgBox "Debe de introducir una variedad destino.", vbExclamation
        Exit Function
    Else
        If Check1(0).Value = 0 And Check1(1).Value = 0 Then
            MsgBox "Debe seleccionar Calibres, Calidades o ambas", vbExclamation
            Exit Function
        End If
    End If
      
    'Llegados aqui OK
    DatosOk = True
        
End Function


Private Function ActualizarRegistros() As Boolean
Dim sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eActualizarRegistros

    ActualizarRegistros = False

    If Check1(0).Value Then ' calibres
        If BloqueaRegistro("calibres", "codvarie = " & DBSet(txtCodigo(70).Text, "N")) Then
            conn.BeginTrans
            If Opcion(0).Value Then ' copiar
                sql = "select * from calibres where codvarie = " & DBSet(NumCod, "N")
                
                Set Rs = New ADODB.Recordset
                Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Not Rs.EOF Then Rs.MoveFirst
                While Not Rs.EOF
                    Sql2 = "select count(*) from calibres where codvarie = " & DBSet(txtCodigo(70).Text, "N")
                    Sql2 = Sql2 & " and codcalib = " & DBSet(Rs!codcalib, "N")
                    
                    If TotalRegistros(Sql2) > 0 Then
                        ' updateamos
                        Sql3 = "update calibres fuente, calibres destino set destino.nomcalib = fuente.nomcalib, "
                        Sql3 = Sql3 & " destino.nomcalab = fuente.nomcalab, destino.calbaneco = fuente.calbaneco "
                        Sql3 = Sql3 & " where fuente.codvarie = " & DBSet(NumCod, "N")
                        Sql3 = Sql3 & " and destino.codvarie = " & DBSet(txtCodigo(70).Text, "N")
                        Sql3 = Sql3 & " and fuente.codcalib = " & DBSet(Rs!codcalib, "N")
                        Sql3 = Sql3 & " and destino.codcalib = " & DBSet(Rs!codcalib, "N")
                
                        conn.Execute Sql3
                    Else
                        ' insertamos
                        Sql3 = "insert into calibres select " & DBSet(txtCodigo(70).Text, "N")
                        Sql3 = Sql3 & ",codcalib, nomcalib, nomcalab, calbaneco from calibres "
                        Sql3 = Sql3 & " where codvarie = " & DBSet(NumCod, "N")
                        Sql3 = Sql3 & " and codcalib = " & DBSet(Rs!codcalib, "N")
        
                        conn.Execute Sql3
                    End If
                    
                    Rs.MoveNext
                Wend
                
                Set Rs = Nothing
                
'                Sql = "delete from calibres where codvarie = " & DBSet(txtCodigo(70).Text, "N")
'                Conn.Execute Sql
'
'                Sql = "insert into calibres select " & DBSet(txtCodigo(70).Text, "N")
'                Sql = Sql & ",codcalib, nomcalib, nomcalab, calbaneco from calibres "
'                Sql = Sql & " where codvarie = " & DBSet(NumCod, "N")
'
'                Conn.Execute Sql
            Else
                sql = "update calibres fuente, calibres destino set destino.nomcalib = fuente.nomcalib, "
                sql = sql & " destino.nomcalab = fuente.nomcalab, destino.calbaneco = fuente.calbaneco "
                sql = sql & " where fuente.codvarie = " & DBSet(NumCod, "N")
                sql = sql & " and destino.codvarie = " & DBSet(txtCodigo(70).Text, "N")
                sql = sql & " and fuente.codcalib = destino.codcalib "
                
                conn.Execute sql
            End If
            conn.CommitTrans
        End If
        TerminaBloquear
    End If
    
    If Check1(1).Value Then ' calidades
        If BloqueaRegistro("rcalidad", "codvarie = " & DBSet(txtCodigo(70).Text, "N")) Then
            conn.BeginTrans
            If Opcion(0).Value Then ' copiar
                sql = "select * from rcalidad where codvarie = " & DBSet(NumCod, "N")
                
                Set Rs = New ADODB.Recordset
                Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Not Rs.EOF Then Rs.MoveFirst
                While Not Rs.EOF
                    Sql2 = "select count(*) from rcalidad where codvarie = " & DBSet(txtCodigo(70).Text, "N")
                    Sql2 = Sql2 & " and codcalid = " & DBSet(Rs!codcalid, "N")
                    
                    If TotalRegistros(Sql2) > 0 Then
                        ' actualizamos
                        Sql3 = "update rcalidad fuente, rcalidad destino set destino.nomcalid = fuente.nomcalid, "
                        Sql3 = Sql3 & " destino.nomcalab = fuente.nomcalab, destino.tipcalid = fuente.tipcalid, "
                        Sql3 = Sql3 & " destino.tipcalid1 = fuente.tipcalid1, "
                        Sql3 = Sql3 & " destino.nomcalibrador1 = fuente.nomcalibrador1, "
                        Sql3 = Sql3 & " destino.nomcalibrador2 = fuente.nomcalibrador2, "
                        Sql3 = Sql3 & " destino.gastosrec = fuente.gastosrec "
                        Sql3 = Sql3 & " where fuente.codvarie = " & DBSet(NumCod, "N")
                        Sql3 = Sql3 & " and destino.codvarie = " & DBSet(txtCodigo(70).Text, "N")
                        Sql3 = Sql3 & " and fuente.codcalid = " & DBSet(Rs!codcalid, "N")
                        Sql3 = Sql3 & " and destino.codcalid = " & DBSet(Rs!codcalid, "N")
                        
                        conn.Execute Sql3
                        
                    Else
                        ' copiamos
                        Sql3 = "insert into rcalidad select " & DBSet(txtCodigo(70).Text, "N")
                        Sql3 = Sql3 & ",codcalid, nomcalid, nomcalab, tipcalid, tipcalid1, nomcalibrador1,"
                        Sql3 = Sql3 & "nomcalibrador2, gastosrec from rcalidad "
                        Sql3 = Sql3 & " where codvarie = " & DBSet(NumCod, "N")
                        Sql3 = Sql3 & " and codcalid = " & DBSet(Rs!codcalid, "N")
                    
                        conn.Execute Sql3
                        
                    End If
                    Rs.MoveNext
                Wend
                Set Rs = Nothing


'                Sql = "delete from rcalidad where codvarie = " & DBSet(txtCodigo(70).Text, "N")
'                Conn.Execute Sql
'
'                Sql = "insert into rcalidad select " & DBSet(txtCodigo(70).Text, "N")
'                Sql = Sql & ",codcalid, nomcalid, nomcalab, tipcalid, tipcalid1, nomcalibrador1,"
'                Sql = Sql & "nomcalibrador2, gastosrec from rcalidad "
'                Sql = Sql & " where codvarie = " & DBSet(NumCod, "N")
'
'                Conn.Execute Sql


            Else
                sql = "update rcalidad fuente, rcalidad destino set destino.nomcalid = fuente.nomcalid, "
                sql = sql & " destino.nomcalab = fuente.nomcalab, destino.tipcalid = fuente.tipcalid, "
                sql = sql & " destino.tipcalid1 = fuente.tipcalid1, "
                sql = sql & " destino.nomcalibrador1 = fuente.nomcalibrador1, "
                sql = sql & " destino.nomcalibrador2 = fuente.nomcalibrador2, "
                sql = sql & " destino.gastosrec = fuente.gastosrec "
                sql = sql & " where fuente.codvarie = " & DBSet(NumCod, "N")
                sql = sql & " and destino.codvarie = " & DBSet(txtCodigo(70).Text, "N")
                sql = sql & " and fuente.codcalid = destino.codcalid "
                
                conn.Execute sql
            End If
            conn.CommitTrans
        End If
        TerminaBloquear
    End If

    ActualizarRegistros = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizar Registros", Err.Description
    conn.RollbackTrans
    TerminaBloquear
End Function


Private Function CalcularDatosFacturaVenta(cadWHERE As String) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim Rs As ADODB.Recordset
Dim i As Integer

Dim sql As String
Dim cadAux As String
Dim cadAux1 As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotNeto As Currency
Dim TotImpIVA As Currency

Dim ImpAux As Currency
Dim impiva As Currency
Dim ImpREC As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA

Dim vBruto As Currency
Dim vNeto As Currency

Dim exentoIVA As Boolean
Dim conDesplaz As Boolean
    
Dim BaseImp As Currency
Dim BaseIVA1 As Currency
Dim BaseIVA2 As Currency
Dim BaseIVA3 As Currency
    
Dim BrutoFac As Currency
    
Dim ImpIVA1 As Currency
Dim ImpIVA2 As Currency
Dim ImpIVA3 As Currency
    
Dim PorceIVA1 As Currency
Dim PorceIVA2 As Currency
Dim PorceIVA3 As Currency
    
Dim ImpREC1 As Currency
Dim ImpREC2 As Currency
Dim ImpREC3 As Currency
    
Dim PorceREC1 As Currency
Dim PorceREC2 As Currency
Dim PorceREC3 As Currency
    
Dim TipoIVA1 As Currency
Dim TipoIVA2 As Currency
Dim TipoIVA3 As Currency
    
Dim ImpDto1 As Currency
Dim ImpDto2 As Currency
Dim TotalFac As Currency

Dim IvaAnt As Integer
Dim cadwhere1 As String
    
Dim Nulo2 As String
Dim Nulo3 As String
Dim TipIvaC As Integer

    CalcularDatosFacturaVenta = False
    On Error GoTo ECalcular

    BaseImp = 0
    BaseIVA1 = 0
    BaseIVA2 = 0
    BaseIVA3 = 0
    
    BrutoFac = 0
    
    ImpIVA1 = 0
    ImpIVA2 = 0
    ImpIVA3 = 0
    
    PorceIVA1 = 0
    PorceIVA2 = 0
    PorceIVA3 = 0
    
    ImpREC1 = 0
    ImpREC2 = 0
    ImpREC3 = 0
    
    PorceREC1 = 0
    PorceREC2 = 0
    PorceREC3 = 0
    
    TipoIVA1 = 0
    TipoIVA2 = 0
    TipoIVA3 = 0
    
    ImpDto1 = 0
    ImpDto2 = 0
    TotalFac = 0

    'Agrupar el importe bruto por tipos de iva
    cadwhere1 = Replace(cadWHERE, "facturas", "facturas_variedad")
    sql = "SELECT facturas_variedad.codigiva, sum(imporbru) as bruto, sum(impornet) as neto"
    sql = sql & " FROM facturas_variedad "
    sql = sql & " WHERE " & cadwhere1
    sql = sql & " GROUP BY 1 "
    sql = sql & " UNION "
    cadwhere1 = Replace(cadWHERE, "facturas", "facturas_envases")
    sql = sql & "SELECT facturas_envases.codigiva, sum(importel) as bruto, sum(importel) as neto"
    sql = sql & " FROM facturas_envases "
    sql = sql & " WHERE " & cadwhere1
    sql = sql & " GROUP BY 1 "
    sql = sql & " UNION "
    cadwhere1 = Replace(cadWHERE, "facturas", "facturas_acuenta")
    sql = sql & "SELECT facturas.codiiva1 as codigiva, sum(brutofac * (-1)) as bruto, sum(brutofac * (-1)) as neto"
    sql = sql & " FROM facturas_acuenta, facturas "
    sql = sql & " WHERE " & cadwhere1
    sql = sql & " and facturas.codtipom = facturas_acuenta.codtipomcta "
    sql = sql & " and facturas.numfactu = facturas_acuenta.numfactucta "
    sql = sql & " and facturas.fecfactu = facturas_acuenta.fecfactucta "
    sql = sql & " GROUP BY 1 "
    sql = sql & " ORDER BY 1 "

    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    TotBruto = 0
    TotNeto = 0
    TotImpIVA = 0
    vBruto = 0
    vNeto = 0
    i = 1

    If Not Rs.EOF Then Rs.MoveFirst
    IvaAnt = Rs.Fields(0).Value
    While Not Rs.EOF
        
        If IvaAnt <> Rs.Fields(0).Value Then
            TotBruto = TotBruto + vBruto
            TotNeto = TotNeto + vNeto
            ImpBImIVA = vNeto
        

            'Obtener el % de IVA
            cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(IvaAnt), "N")

            'aplicar el IVA a la base imponible de ese tipo
            impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
            
            'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
            'los vamos acumulando
            TotImpIVA = TotImpIVA + impiva

            TipIvaC = DevuelveValor("select tipoivac from facturas where " & cadWHERE)
 

            If CInt(TipIvaC) = 2 Then
                'Obtener el % de RECARGO
                cadAux1 = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", CStr(IvaAnt), "N")
    
                'aplicar el RECARGO a la base imponible de ese tipo
                ImpREC = CalcularPorcentaje(ImpBImIVA, CCur(cadAux1), 2)
                
                'sumamos todos los RECARGOS para sumarselo a la base imponible total de la factura
                'los vamos acumulando
                TotImpIVA = TotImpIVA + ImpREC
            Else
                cadAux1 = "0"
                ImpREC = 0
            End If


            Select Case i
                Case 1  'IVA 1
                    TipoIVA1 = IvaAnt 'RS!codigiva

                    BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA1 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA1 = impiva
                    
                    PorceREC1 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC1 = ImpREC

                Case 2  'IVA 2
                    TipoIVA2 = IvaAnt 'RS!codigiva

                    BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA2 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA2 = impiva

                    PorceREC2 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC2 = ImpREC
                Case 3  'IVA 3
                    TipoIVA3 = IvaAnt 'RS!codigiva

                    BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA3 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA3 = impiva
                    
                    PorceREC3 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC3 = ImpREC
            End Select
            
            
            i = i + 1
            IvaAnt = Rs.Fields(0).Value
            vBruto = DBLet(Rs.Fields(1).Value, "N")
            vNeto = DBLet(Rs.Fields(2).Value, "N")
        Else
            vBruto = vBruto + DBLet(Rs.Fields(1).Value, "N")
            vNeto = vNeto + DBLet(Rs.Fields(2).Value, "N")
        End If
        
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    ' ULTIMO REGISTRO
    TotBruto = TotBruto + vBruto
    TotNeto = TotNeto + vNeto
    ImpBImIVA = vNeto


    'Obtener el % de IVA
    cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(IvaAnt), "N")

    'aplicar el IVA a la base imponible de ese tipo
    impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
    
    'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
    'los vamos acumulando
    TotImpIVA = TotImpIVA + impiva
    
    TipIvaC = DevuelveValor("select tipoivac from facturas where " & cadWHERE)
   
    If CInt(TipIvaC) = 2 Then
        'Obtener el % de RECARGO
        cadAux1 = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", CStr(IvaAnt), "N")
    
        'aplicar el RECARGO a la base imponible de ese tipo
        ImpREC = CalcularPorcentaje(ImpBImIVA, CCur(cadAux1), 2)
    Else
        cadAux1 = "0"
        ImpREC = 0
    End If
    'sumamos todos los RECARGOS para sumarselo a la base imponible total de la factura
    'los vamos acumulando
    TotImpIVA = TotImpIVA + ImpREC



    Select Case i
        Case 1  'IVA 1
            TipoIVA1 = IvaAnt

            BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE

            PorceIVA1 = cadAux '% de IVA

            'Importe total con IVA
            ImpIVA1 = impiva
            
            PorceREC1 = cadAux1 '% de REC

            'Importe total con RECARGO
            ImpREC1 = ImpREC

        Case 2  'IVA 2
            TipoIVA2 = IvaAnt

            BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE

            PorceIVA2 = cadAux '% de IVA

            'Importe total con IVA
            ImpIVA2 = impiva

            PorceREC2 = cadAux1 '% de REC

            'Importe total con RECARGO
            ImpREC2 = ImpREC
        Case 3  'IVA 3
            TipoIVA3 = IvaAnt

            BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE

            PorceIVA3 = cadAux '% de IVA

            'Importe total con IVA
            ImpIVA3 = impiva
            
            PorceREC3 = cadAux1 '% de REC

            'Importe total con RECARGO
            ImpREC3 = ImpREC
    End Select

    'Base Imponible
    BaseImp = TotNeto

    'TOTAL de la factura
    TotalFac = BaseImp + TotImpIVA

    'ACTUALIZAMOS LA FACTURA (tabla facturas)
    sql = "update facturas "
    sql = sql & "set baseimp1 = " & DBSet(BaseIVA1, "N")
    sql = sql & ",impoiva1 = " & DBSet(ImpIVA1, "N")
    sql = sql & ",imporec1 = " & DBSet(ImpREC1, "N")
    sql = sql & ",porciva1 = " & DBSet(PorceIVA1, "N")
    sql = sql & ",porcrec1 = " & DBSet(PorceREC1, "N")
    sql = sql & ",codiiva1 = " & DBSet(TipoIVA1, "N")
    Nulo2 = "N"
    Nulo3 = "N"
    If DBSet(TipoIVA2, "N", "S") = ValorNulo Then Nulo2 = "S"
    If DBSet(TipoIVA3, "N", "S") = ValorNulo Then Nulo3 = "S"
    sql = sql & ",baseimp2 = " & DBSet(BaseIVA2, "N", Nulo2)
    sql = sql & ",impoiva2 = " & DBSet(ImpIVA2, "N", Nulo2)
    sql = sql & ",imporec2 = " & DBSet(ImpREC2, "N", Nulo2)
    sql = sql & ",porciva2 = " & DBSet(PorceIVA2, "N", Nulo2)
    sql = sql & ",porcrec2 = " & DBSet(PorceREC2, "N", Nulo2)
    sql = sql & ",codiiva2 = " & DBSet(TipoIVA2, "N", Nulo2)
    sql = sql & ",baseimp3 = " & DBSet(BaseIVA3, "N", Nulo3)
    sql = sql & ",impoiva3 = " & DBSet(ImpIVA3, "N", Nulo3)
    sql = sql & ",imporec3 = " & DBSet(ImpREC3, "N", Nulo3)
    sql = sql & ",porciva3 = " & DBSet(PorceIVA3, "N", Nulo3)
    sql = sql & ",porcrec3 = " & DBSet(PorceREC3, "N", Nulo3)
    sql = sql & ",codiiva3 = " & DBSet(TipoIVA3, "N", Nulo3)
    sql = sql & ",brutofac = " & DBSet(TotBruto, "N")
    sql = sql & ",impordto = " & DBSet(Round2(TotBruto - TotNeto, 2), "N")
    sql = sql & ",totalfac = " & DBSet(TotalFac, "N")
    sql = sql & " where " & cadWHERE
    
    conn.Execute sql

    CalcularDatosFacturaVenta = True

ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFacturaVenta = False
    Else
        CalcularDatosFacturaVenta = True
    End If
End Function

