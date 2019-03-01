Attribute VB_Name = "ModCostes"
Public Function ActualizarCostes(Albaran As Long, Linea As Integer, Insertar As Boolean, ForfaitAnt As String, CodPaletAnt As String) As Boolean
' Insertar: indica si hemos de volver a generar los costes en albaran_costes
Dim b As Boolean
Dim Forfait As String
Dim CodPalet As String

    On Error GoTo eActualizarCostes
'--monica:antes
'    Forfait = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "codforfait", "numalbar", CStr(albaran), "N", , "numlinea", CStr(Linea), "N")
'    b = BorrarEnvases(albaran, Linea, Forfait)
'++monica:ahora
    b = BorrarEnvases(Albaran, Linea, ForfaitAnt)
    Forfait = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "codforfait", "numalbar", CStr(Albaran), "N", , "numlinea", CStr(Linea), "N")
    
    '[Monica] 15/06/2010 añadido costes paletizacion
    CodPalet = ""
    If CodPaletAnt <> "" Then
        If b Then b = BorrarEnvasesPalet(Albaran, Linea, CodPaletAnt)
        CodPalet = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "codpalet", "numalbar", CStr(Albaran), "N", , "numlinea", CStr(Linea), "N")
    End If
    
    If b Then b = BorrarTablaCostes(Albaran, Linea)
    
    If b And Insertar Then
        b = InsertarTablaCostes(Albaran, Linea, Forfait)
        
        If b Then b = InsertarEnvases(Albaran, Linea, Forfait)
        
        '[Monica] 15/06/2010 añadido costes paletizacion
        If CodPalet <> "" Then
            If b Then b = InsertarEnvasesPalet(Albaran, Linea, CodPalet)
        End If
        
    End If
        
eActualizarCostes:
    If Err.Number <> 0 Or Not b Then
        ActualizarCostes = False
    Else
        ActualizarCostes = True
    End If
End Function


Public Function BorrarTablaCostes(Albaran As Long, Linea As Integer) As Boolean
Dim Sql As String
    
    On Error GoTo EBorrar
    
    BorrarTablaCostes = True
    
    Sql = "delete from albaran_costes where numalbar = " & Albaran & " and numlinea = " & Linea
    
'08/09/2009: modificacion hecha pq se borraban los costes de portes, cuando se revaloraban
'            albaranes
    Sql = Sql & " and albaran_costes.tipogasto <> 2 "
    Sql = Sql & " and albaran_costes.tipogasto <> 3 "
    
    
    conn.Execute Sql
    
    Exit Function

EBorrar:
    BorrarTablaCostes = False
End Function


Public Function InsertarTablaCostes(Albaran As Long, Linea As Integer, Forfait As String) As Boolean
Dim Sql As String
Dim Rs As Recordset
Dim Cajas As String
Dim Kilos As String
Dim Importe As Currency
Dim CajaKilo As Byte '0 caja 1 kilo
                    
    On Error GoTo EInsertarTablaCostes
    
    InsertarTablaCostes = False
    
    Set Rs = New ADODB.Recordset
    
    CajaKilo = DevuelveDesdeBDNew(cAgro, "forfaits", "cajakilo", "codforfait", Forfait, "T")
    
    Sql = "select * from forfaits_costes where codforfait = " & DBSet(Forfait, "T")
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Importe = 0
        
        Sql = "INSERT INTO albaran_costes (numalbar,numlinea,tipogasto,codcoste,impcoste,importes,unidades, codartic)"
        Sql = Sql & " VALUES (" & DBSet(Albaran, "N") & "," & DBSet(Linea, "N") & ",0," & DBSet(Rs!codCoste, "N") & ","
        
        Select Case CajaKilo
            Case 0 ' por caja
                Cajas = ""
                Cajas = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "numcajas", "numalbar", CStr(Albaran), "N", , "numlinea", CStr(Linea), "N")
                
                Importe = Round2(ComprobarCero(Cajas) * DBLet(Rs!importes, "N"), 4)
                
                Sql = Sql & DBSet(Importe, "N") & "," & DBSet(Rs!importes, "N") & "," & DBSet(Cajas, "N") & "," & ValorNulo & ")"

            
            Case 1 ' por kilo
                Kilos = ""
                Kilos = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "pesoneto", "numalbar", CStr(Albaran), "N", , "numlinea", CStr(Linea), "N")
            
                Importe = Round2(ComprobarCero(Kilos) * DBLet(Rs!importes, "N"), 4)
                
                Sql = Sql & DBSet(Importe, "N") & "," & DBSet(Rs!importes, "N") & "," & DBSet(Kilos, "N") & "," & ValorNulo & ")"
        End Select
        
        conn.Execute Sql
    
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    InsertarTablaCostes = True
    Exit Function

EInsertarTablaCostes:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Insertar Tabla de Costes", Err.Description
    End If
End Function

Public Function InsertarEnvases(Albaran As Long, Linea As Integer, Forfait As String) As Boolean
Dim Sql As String
Dim Rs As Recordset
Dim Cajas As String
Dim Kilos As String
Dim Importe As Currency
Dim Precio As Currency
Dim PrecioLin As Currency
Dim vCStock As CStock
Dim MenError As String
Dim b As Boolean
Dim vCajas As String
    On Error GoTo EInsertarEnvases
    
    InsertarEnvases = False
    
    Set Rs = New ADODB.Recordset
    
    Sql = "select codartic, sum(cantidad) from forfaits_envases where codforfait = " & DBSet(Forfait, "T")
    Sql = Sql & " group by 1 order by 1 "
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Cajas = ""
    Cajas = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "numcajas", "numalbar", CStr(Albaran), "N", , "numlinea", CStr(Linea), "N")
    b = True
    
    While Not Rs.EOF And b
        Precio = PrecioEnvase(DBLet(Rs!codArtic, "T"))
        PrecioLin = Round2(Precio * DBLet(Rs.Fields(1).Value, "N"), 4)
        
        Importe = 0
        Importe = Round2(PrecioLin * CCur(DBSet(Cajas, "N")), 4)
        
'[Monica] 27/08/2010 : añadida esta instruccion
'                   me voy a guardar no las cajas sino la cantidad del articulo (smoval.cantidad)
        vCajas = CCur(ComprobarCero(Cajas)) * DBLet(Rs.Fields(1).Value, "N")
         
        
        Sql = "INSERT INTO albaran_costes (numalbar,numlinea,tipogasto,codcoste,impcoste,importes,unidades,codartic)"
        Sql = Sql & "VALUES (" & DBSet(Albaran, "N") & "," & DBSet(Linea, "N") & ",1," & ValorNulo & ","
        Sql = Sql & DBSet(Importe, "N") & "," & DBSet(PrecioLin, "N") & "," & DBSet(vCajas, "N") & "," & DBSet(Rs!codArtic, "T") & ")"
        
        conn.Execute Sql
    
        ' insertamos el movimiento y reducimos el stock
        Set vCStock = New CStock
'        b = InicializarCStock(vCStock, "S", Albaran, Linea, RS!codArtic, CCur(ComprobarCero(Cajas)) * DBLet(RS.Fields(1).Value, "N"), Importe)
'[Monica] 27/08/2010 : cambiada la anterior intruccion por  esta instruccion
'                   me voy a guardar no las cajas sino la cantidad del articulo (smoval.cantidad)
            b = InicializarCStock(vCStock, "S", Albaran, Linea, Rs!codArtic, CCur(ComprobarCero(vCajas)), Importe)

        'en actualizar stock comprobamos si el articulo tiene control de stock
        If b Then
            MenError = "Insertar en almacen"
            b = InsertarAlmacen(vCStock)
            MenError = "Actualizando Stocks"
            b = vCStock.ActualizarStock(True)
        End If
        Set vCStock = Nothing

        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    InsertarEnvases = b
    Exit Function

EInsertarEnvases:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Insertar Tabla de Costes", Err.Description
    End If
End Function


Public Function InsertarEnvasesPalet(Albaran As Long, Linea As Integer, CodPalet As String) As Boolean
Dim Sql As String
Dim Rs As Recordset
Dim Cajas As String
Dim Kilos As String
Dim Importe As Currency
Dim Precio As Currency
Dim PrecioLin As Currency
Dim vCStock As CStock
Dim MenError As String
Dim b As Boolean

Dim TotPalet As String
Dim vTotPalet As String

    On Error GoTo EInsertarEnvases
    
    InsertarEnvasesPalet = False
    
    Set Rs = New ADODB.Recordset
    
    Sql = "select totpalet from albaran_variedad where numalbar = " & DBSet(Albaran, "N")
    Sql = Sql & " and numlinea = " & DBSet(Linea, "N")
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    TotPalet = DBLet(Rs!TotPalet, "N")
    
    Set Rs = Nothing
    
    Set Rs = New ADODB.Recordset
    
    Sql = "select codartic, sum(cantidad) from confpale_envases where codpalet = " & DBSet(CodPalet, "N")
    Sql = Sql & " group by 1 order by 1 "
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    b = True
    
    While Not Rs.EOF And b
        Precio = PrecioEnvase(DBLet(Rs!codArtic, "T"))
        PrecioLin = Round2(Precio * DBLet(Rs.Fields(1).Value, "N"), 4)
        
        Importe = 0
        Importe = Round2(PrecioLin * CCur(DBSet(TotPalet, "N")), 4)

'[Monica] 27/08/2010 : añadida esta instruccion
'                   me voy a guardar no las cajas sino la cantidad del articulo (smoval.cantidad)
        vTotPalet = CStr(CCur(ComprobarCero(TotPalet)) * DBLet(Rs.Fields(1).Value, "N"))
        
        Sql = "INSERT INTO albaran_costes (numalbar,numlinea,tipogasto,codcoste,impcoste,importes,unidades,codartic)"
        Sql = Sql & "VALUES (" & DBSet(Albaran, "N") & "," & DBSet(Linea, "N") & ",4," & ValorNulo & ","
        Sql = Sql & DBSet(Importe, "N") & "," & DBSet(PrecioLin, "N") & "," & DBSet(vTotPalet, "N") & "," & DBSet(Rs!codArtic, "T") & ")"
        
        conn.Execute Sql
    
        ' insertamos el movimiento y reducimos el stock
        Set vCStock = New CStock
'        b = InicializarCStock(vCStock, "S", Albaran, Linea, RS!codArtic, CCur(ComprobarCero(TotPalet)) * DBLet(RS.Fields(1).Value, "N"), Importe)
'[Monica] 27/08/2010 : cambiada la anterior intruccion por  esta instruccion
'                   me voy a guardar no las cajas sino la cantidad del articulo (smoval.cantidad)
            b = InicializarCStock(vCStock, "S", Albaran, Linea, Rs!codArtic, CCur(ComprobarCero(vTotPalet)), Importe)
    
        'en actualizar stock comprobamos si el articulo tiene control de stock
        If b Then
            MenError = "Insertar en almacen"
            b = InsertarAlmacen(vCStock)
            MenError = "Actualizando Stocks"
            b = vCStock.ActualizarStock(True)
        End If
        Set vCStock = Nothing

        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    InsertarEnvasesPalet = b
    Exit Function

EInsertarEnvases:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Insertar Tabla de Costes", Err.Description
    End If
End Function




Public Function BorrarEnvases(Albaran As Long, Linea As Integer, Forfait As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As Recordset
Dim Rs2 As Recordset
Dim Cajas As String
Dim Kilos As String
Dim Importe As Currency
Dim Precio As Currency
Dim vCStock As CStock
Dim MenError As String

    On Error GoTo EBorrarEnvases
    
    BorrarEnvases = False
    
    Set Rs = New ADODB.Recordset
    
    Sql = "select * from albaran_costes where numalbar = " & DBSet(Albaran, "N")
    Sql = Sql & " and numlinea = " & DBSet(Linea, "N") & " and tipogasto = 1 "
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF
        Sql2 = "select sum(cantidad) from forfaits_envases where codforfait = " & DBSet(Forfait, "T")
        Sql2 = Sql2 & " and codartic = " & DBSet(Rs!codArtic, "T")
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Cajas = 0
'        If Not Rs2.EOF Then Cajas = DBLet(Rs2.Fields(0).Value, "N")
'[Monica] 27/08/2010 : cambiado lo anterior por la siguiente instruccion
'                   me voy a guardar no las cajas sino la cantidad del articulo (smoval.cantidad)
        Cajas = DBLet(Rs!Unidades, "N")
        
        Set Rs2 = Nothing
        
        ' borramos el movimiento y aumentamos el stock
        Set vCStock = New CStock
'        If Not InicializarCStock(vCStock, "E", Albaran, Linea, RS!codArtic, DBLet(RS!Unidades, "N") * DBLet(Cajas, "N"), DBLet(RS!importes, "N")) Then Exit Function
'[Monica] 27/08/2010 : cambiado lo anterior por la siguiente instruccion
'                   me voy a guardar no las cajas sino la cantidad del articulo (smoval.cantidad)
        If Not InicializarCStock(vCStock, "E", Albaran, Linea, Rs!codArtic, DBLet(Cajas, "N"), DBLet(Rs!importes, "N")) Then Exit Function
   
        'en actualizar stock comprobamos si el articulo tiene control de stock
        b = vCStock.DevolverStock
        Set vCStock = Nothing
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    BorrarEnvases = True
    Exit Function

EBorrarEnvases:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Borrar Tabla de Costes", Err.Description
    End If
End Function

'[Monica] 15/06/2010 añadido costes paletizacion
Public Function BorrarEnvasesPalet(Albaran As Long, Linea As Integer, CodPalet As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As Recordset
Dim Rs2 As Recordset
Dim Cantidad As String
Dim Kilos As String
Dim Importe As Currency
Dim Precio As Currency
Dim vCStock As CStock
Dim MenError As String


    On Error GoTo EBorrarEnvases
    
    BorrarEnvasesPalet = False
    
    Set Rs = New ADODB.Recordset
    
    Sql = "select * from albaran_costes where numalbar = " & DBSet(Albaran, "N")
    Sql = Sql & " and numlinea = " & DBSet(Linea, "N") & " and tipogasto = 4 "
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF
        Sql2 = "select sum(cantidad) from confpale_envases where codpalet = " & DBSet(CodPalet, "N")
        Sql2 = Sql2 & " and codartic = " & DBSet(Rs!codArtic, "T")
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Cantidad = 0
'        If Not Rs2.EOF Then cantidad = DBLet(Rs2.Fields(0).Value, "N")
'[Monica] 27/08/2010 : cambiado lo anterior por la siguiente instruccion
'                   me voy a guardar no las cajas sino la cantidad del articulo (smoval.cantidad)
        Cantidad = DBLet(Rs!Unidades, "N")
        
        Set Rs2 = Nothing
    
    
        ' borramos el movimiento y aumentamos el stock
        Set vCStock = New CStock
'        If Not InicializarCStock(vCStock, "E", Albaran, Linea, RS!codArtic, DBLet(RS!Unidades, "N") * DBLet(cantidad, "N"), DBLet(RS!importes, "N")) Then Exit Function
'[Monica] 27/08/2010 : cambiado lo anterior por la siguiente instruccion
'                   me voy a guardar no las cajas sino la cantidad del articulo (smoval.cantidad)
        If Not InicializarCStock(vCStock, "E", Albaran, Linea, Rs!codArtic, DBLet(Cantidad, "N"), DBLet(Rs!importes, "N")) Then Exit Function
   
        'en actualizar stock comprobamos si el articulo tiene control de stock
        b = vCStock.DevolverStock
        Set vCStock = Nothing
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    BorrarEnvasesPalet = True
    Exit Function

EBorrarEnvases:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Borrar Tabla de Costes", Err.Description
    End If
End Function

'Private Sub CalcularPrecio()
'    txtAux2(1).Text = ""
'    If txtAux(2).Text <> "" And txtAux2(0).Text <> "" Then
'        txtAux2(1).Text = Round2(ImporteSinFormato(txtAux(2).Text) * ImporteSinFormato(txtAux2(0).Text), 4)
'    End If
'End Sub

Private Function PrecioEnvase(Artic As String) As Currency
    PrecioEnvase = 0
    
    If Artic = "" Then Exit Function
    
    Select Case vParamAplic.TipoPrecio
        Case 0
            PrecioEnvase = DevuelveDesdeBDNew(cAgro, "sartic", "preciomp", "codartic", Artic, "T")
        Case 1
            PrecioEnvase = DevuelveDesdeBDNew(cAgro, "sartic", "preciouc", "codartic", Artic, "T")
    End Select
End Function

Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Albaran As Long, Linea As Integer, Artic As String, Cantidad As Currency, Importe As Currency) As Boolean
'On Error Resume Next
On Error Resume Next

    vCStock.tipoMov = TipoM 'Movimiento de Entrada o Salida
    vCStock.DetaMov = vParamAplic.CodTipomAlb ' "ALV" '"ALC=Albaran de Compra"
    vCStock.Fechamov = DevuelveDesdeBDNew(cAgro, "albaran", "fechaalb", "numalbar", CStr(Albaran), "N")
    vCStock.Trabajador = DevuelveDesdeBDNew(cAgro, "albaran", "codclien", "numalbar", CStr(Albaran), "N") 'En smoval guardamos el Proveedor
    vCStock.Documento = Albaran
    
    vCStock.codArtic = Artic
    vCStock.codAlmac = DevuelveDesdeBDNew(cAgro, "albaran", "codalmac", "numalbar", CStr(Albaran), "N")
    vCStock.Cantidad = CSng(Cantidad)
    vCStock.Importe = CCur(Importe)
    vCStock.LineaDocu = Linea
    
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function

Public Function TotalCostesEnvases(Albaran As Long, Linea As Integer, Tipo As Byte) As String
Dim Rs As ADODB.Recordset
Dim Sql As String

    Set Rs = New ADODB.Recordset
    
    Sql = "select sum(impcoste) from albaran_costes where numalbar = " & DBSet(Albaran, "N")
    If Linea <> -1 Then
        Sql = Sql & " and numlinea = " & DBSet(Linea, "N")
    End If
    Sql = Sql & " and tipogasto = " & DBSet(Tipo, "N")
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Sql = "0"
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            Sql = CStr(Rs.Fields(0))
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    TotalCostesEnvases = Sql
    

End Function


Public Function EliminarCostes(Albaran As Long) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
    
    On Error GoTo eEliminarCostes
    
    Set Rs = New ADODB.Recordset
    
    Sql = "select * from albaran_variedad where numalbar = " & DBSet(Albaran, "N")
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    b = True
    While Not Rs.EOF And b
        b = ActualizarCostes(Albaran, DBLet(Rs!NumLinea, "N"), False, DBLet(Rs!codforfait, "T"), DBLet(Rs!CodPalet, "N"))
        Rs.MoveNext
    Wend

    Set Rs = Nothing
eEliminarCostes:
    If Err.Number <> 0 Or Not b Then
        EliminarCostes = False
    Else
        EliminarCostes = True
    End If
        
End Function


Public Function InsertarCostes(Albaran As Long) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
    
    On Error GoTo eInsertarCostes
    
    Set Rs = New ADODB.Recordset
    
    Sql = "select * from albaran_variedad where numalbar = " & DBSet(Albaran, "N")
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    b = True
    While Not Rs.EOF And b
        b = ActualizarCostes(Albaran, DBLet(Rs!NumLinea, "N"), True, DBLet(Rs!codforfait, "T"), DBLet(Rs!CodPalet, "N"))
        Rs.MoveNext
    Wend

    Set Rs = Nothing
eInsertarCostes:
    If Err.Number <> 0 Or Not b Then
        InsertarCostes = False
    Else
        InsertarCostes = True
    End If
End Function



Private Function InsertarAlmacen(ByRef vCStock As CStock) As Boolean
'Inserta en la tabla salmac si no está el registro
Dim cadMen As String
Dim CadValues As String 'cadena para la SQL de insertar en la tabla salmac

    On Error GoTo EInsertarAlmacen
    
    'comprobar que el articulo esta en almacen
    cadMen = ""
    cadMen = DevuelveDesdeBDNew(cAgro, "salmac", "codartic", "codartic", vCStock.codArtic, "T", , "codalmac", vCStock.codAlmac, "N")
    If cadMen = "" Then 'se tiene que insertar
        CadValues = "INSERT INTO salmac (codartic,codalmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin)"
        CadValues = CadValues & " VALUES (" & DBSet(vCStock.codArtic, "T") & "," & DBSet(vCStock.codAlmac, "N") & ",0,0,0,0,0,NULL,NULL,0)"
        
        conn.Execute CadValues
    End If
    InsertarAlmacen = True
    Exit Function
    
EInsertarAlmacen:
    If Err.Number <> 0 Then
        InsertarAlmacen = False
    End If
End Function

Public Function FacturaContabilizada(Albaran As Currency, Linea As Currency) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim Total As Currency
Dim vHayReg As Currency

    On Error GoTo eFacturaContabilizada

    FacturaContabilizada = 0
    
    Sql = "select count(*) from facturas, facturas_variedad "
    Sql = Sql & " where facturas_variedad.numalbar = " & Albaran
    Sql = Sql & " and facturas_variedad.numlinealbar = " & Linea
    Sql = Sql & " and facturas_variedad.codtipom = facturas.codtipom "
    Sql = Sql & " and facturas_variedad.numfactu = facturas.numfactu "
    Sql = Sql & " and facturas_variedad.fecfactu = facturas.fecfactu "
    Sql = Sql & " and facturas.intconta = 1"
    
    FacturaContabilizada = (TotalRegistros(Sql) > 0)

eFacturaContabilizada:
    If Err.Number <> 0 Then
        FacturaCntabilizada = 0
    End If
End Function



Public Function FacturaCobradaTesoreria(Albaran As Currency, Linea As Currency) As Byte
Dim Sql As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim Total As Currency
Dim vHayReg As Currency

    On Error GoTo eFacturaCobradaTesoreria

    FacturaCobradaTesoreria = 0

    ' seleccionamos las facturas en donde aparece el albaran-linea
    Sql = "select distinct stipom.letraser, facturas_variedad.numfactu, facturas_variedad.fecfactu"
    Sql = Sql & " from facturas_variedad, usuarios.stipom stipom  "
    Sql = Sql & " where facturas_variedad.numalbar = " & Albaran
    Sql = Sql & " and facturas_variedad.numlinealbar = " & Linea
    Sql = Sql & " and facturas_variedad.codtipom = stipom.codtipom "
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Total = 0
    While Not Rs.EOF
        If vParamAplic.ContabilidadNueva Then
            Sql = "select sum(if(isnull(impvenci),0,impvenci) - if(isnull(impcobro),0,impcobro)) from cobros where numserie = " & DBSet(Rs.Fields(0).Value, "T")
            Sql = Sql & " and numfactu = " & DBSet(Rs.Fields(1).Value, "N")
            Sql = Sql & " and fecfactu = " & DBSet(Rs.Fields(2).Value, "F")
        Else
            Sql = "select sum(if(isnull(impvenci),0,impvenci) - if(isnull(impcobro),0,impcobro)) from scobro where numserie = " & DBSet(Rs.Fields(0).Value, "T")
            Sql = Sql & " and codfaccl = " & DBSet(Rs.Fields(1).Value, "N")
            Sql = Sql & " and fecfaccl = " & DBSet(Rs.Fields(2).Value, "F")
        End If
        Set Rs1 = New ADODB.Recordset
        Rs1.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs1.EOF And Not IsNull(Rs1.Fields(0)) Then
            Total = Total + DBLet(Rs1.Fields(0).Value, "N")
'++monica:10/02/2009 si me devuelve nulo no hay nada en la scobro
        Else
'            Exit Function
'++
        End If
    
        Rs.MoveNext
    Wend
    
    If Total = 0 Then
        FacturaCobradaTesoreria = 1
    End If
    Exit Function
    
eFacturaCobradaTesoreria:
    If Err.Number <> 0 Then
        FacturaCobradaTesoreria = 0
    End If
End Function



Public Function AlbaranCobradoTesoreria(Albaran As Currency, Linea As Currency) As Byte
Dim Sql As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim Total As Currency

Dim Cliente As String

    On Error GoTo eAlbaranCobradoTesoreria

    AlbaranCobradoTesoreria = 1

    If Not FacturaContabilizada(Albaran, Linea) Then
        AlbaranCobradoTesoreria = 0
        Exit Function
    End If

    If EsClienteConCtrolCobroAlbaran(Albaran, Linea) Then
        If vParamAplic.ContabilidadNueva Then
            Sql = "select sum(if(isnull(impvenci),0,impvenci) - if(isnull(impcobro),0,impcobro)) from cobros where referencia1 = " & DBSet(Albaran, "N")
            Sql = Sql & " and referencia2 = " & DBSet(Linea, "N")
        Else
            Sql = "select sum(if(isnull(impvenci),0,impvenci) - if(isnull(impcobro),0,impcobro)) from scobro where referencia1 = " & DBSet(Albaran, "N")
            Sql = Sql & " and referencia2 = " & DBSet(Linea, "N")
        End If
        Set Rs1 = New ADODB.Recordset
        Rs1.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs1.EOF And Not IsNull(Rs1.Fields(0)) Then
            If Rs1.Fields(0).Value <> 0 Then AlbaranCobradoTesoreria = 0
        End If
   
    Else
        AlbaranCobradoTesoreria = FacturaCobradaTesoreria(Albaran, Linea)
    End If
   
eAlbaranCobradoTesoreria:
    If Err.Number <> 0 Then
        AlbaranCobradoTesoreria = 0
    End If
End Function


Public Function EsClienteConCtrolCobroAlbaran(Albaran As Currency, Linea As Currency) As Boolean

Dim Sql As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset

    On Error GoTo eEsClienteConCtrolCobroAlbaran
    
    EsClienteConCtrolCobroAlbaran = False
    
    Sql = "select clientes.ctrolcobroalb from (albaran inner join albaran_variedad On albaran_variedad.numalbar = albaran.numalbar) "
    Sql = Sql & " INNER JOIN clientes On albaran.codclien = clientes.codclien "
    Sql = Sql & " where albaran_variedad.numalbar = "
    Sql = Sql & DBSet(Albaran, "N") & " and albaran_variedad.numlinea = " & DBSet(Linea, "N")
    
    EsClienteConCtrolCobroAlbaran = (DevuelveValor(Sql) = 1)
    
eEsClienteConCtrolCobroAlbaran:
    If Err.Number <> 0 Then
        EsClienteConCtrolCobroAlbaran = False
    End If
End Function



Public Function AlbaranFacturado(Albaran As Currency, Linea As Currency) As Byte
Dim Sql As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim Total As Currency

    On Error GoTo eAlbaranFacturado

    AlbaranFacturado = 0

    ' seleccionamos las facturas en donde aparece el albaran-linea
    Sql = "select count(*) "
    Sql = Sql & " from facturas_variedad "
    Sql = Sql & " where facturas_variedad.numalbar = " & Albaran
    Sql = Sql & " and facturas_variedad.numlinealbar = " & Linea
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0).Value, "N") > 0 Then AlbaranFacturado = 1
    End If
    
    Exit Function
eAlbaranFacturado:
    If Err.Number <> 0 Then
        AlbaranFacturado = 0
    End If
End Function

Public Function ImporteAlbaranFacturado(Albaran As Currency, Linea As Currency) As Double
Dim Sql As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim Total As Currency

    On Error GoTo eImporteAlbaranFacturado

    ImporteAlbaranFacturado = 0

    ' seleccionamos las facturas en donde aparece el albaran-linea
    Sql = "select sum(impornet) "
    Sql = Sql & " from facturas_variedad "
    Sql = Sql & " where facturas_variedad.numalbar = " & Albaran
    Sql = Sql & " and facturas_variedad.numlinealbar = " & Linea
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        ImporteAlbaranFacturado = DBLet(Rs.Fields(0).Value, "N")
    End If
    
    Exit Function
eImporteAlbaranFacturado:
    If Err.Number <> 0 Then
        ImporteAlbaranFacturado = 0
    End If
End Function

Public Function FacturasdeAlbaran(Albaran As Currency, Linea As Currency) As String
Dim Sql As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim Total As Currency

    On Error GoTo eFacturasdeAlbaran

    FacturasdeAlbaran = ""

    ' seleccionamos las facturas en donde aparece el albaran-linea
    Sql = "select codtipom, numfactu, fecfactu "
    Sql = Sql & " from facturas_variedad "
    Sql = Sql & " where facturas_variedad.numalbar = " & Albaran
    Sql = Sql & " and facturas_variedad.numlinealbar = " & Linea
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
        Cad = Cad & "(" & DBSet(Rs.Fields(0).Value, "T") & "," & DBSet(Rs.Fields(1).Value, "N") & "," & DBSet(Rs.Fields(2).Value, "F") & "),"
        
        Rs.MoveNext
    Wend
    ' quitamos la ultima ,
    If Cad <> "" Then Cad = Mid(Cad, 1, Len(Cad) - 1)
    
    '[Monica]05/03/2013: indicamos la tabla pq sino da error en el vista previa
    FacturasdeAlbaran = "(facturas.codtipom, facturas.numfactu, facturas.fecfactu) in (" & Cad & ")"
    Exit Function
    
eFacturasdeAlbaran:
    If Err.Number <> 0 Then
        FacturasdeAlbaran = ""
    End If
End Function

'####################################################################################
'##############   FACTURAS DE VENTA A SOCIOS DE RECOLECCION   #######################
'####################################################################################

Public Function AlbaranSOCIO_CobradoTesoreria(Albaran As Currency, Linea As Currency) As Byte
Dim Sql As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim Total As Currency

Dim Cliente As String

    On Error GoTo eAlbaranSOCIO_CobradoTesoreria

    AlbaranSOCIO_CobradoTesoreria = 1

    If EsClienteConCtrolCobroAlbaran(Albaran, Linea) Then

        If vParamAplic.ContabilidadNueva Then
            Sql = "select sum(if(isnull(impvenci),0,impvenci) - if(isnull(impcobro),0,impcobro)) from cobros where referencia1 = " & DBSet(Albaran, "N")
            Sql = Sql & " and referencia2 = " & DBSet(Linea, "N")
        Else
            Sql = "select sum(if(isnull(impvenci),0,impvenci) - if(isnull(impcobro),0,impcobro)) from scobro where referencia1 = " & DBSet(Albaran, "N")
            Sql = Sql & " and referencia2 = " & DBSet(Linea, "N")
        End If
        Set Rs1 = New ADODB.Recordset
        Rs1.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs1.EOF And Not IsNull(Rs1.Fields(0)) Then
            If Rs1.Fields(0).Value <> 0 Then AlbaranSOCIO_CobradoTesoreria = 0
        End If
   
    Else
        AlbaranSOCIO_CobradoTesoreria = FacturaSOCIO_CobradaTesoreria(Albaran, Linea)
    End If
   
eAlbaranSOCIO_CobradoTesoreria:
    If Err.Number <> 0 Then
        AlbaranSOCIO_CobradoTesoreria = 0
    End If
End Function

Public Function FacturaSOCIO_CobradaTesoreria(Albaran As Currency, Linea As Currency) As Byte
Dim Sql As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim Total As Currency

    On Error GoTo eFacturaSOCIO_CobradaTesoreria

    FacturaSOCIO_CobradaTesoreria = 0

    ' seleccionamos las facturas en donde aparece el albaran-linea
    Sql = "select distinct stipom.letraser, facturas_variedad.numfactu, facturas_variedad.fecfactu"
    Sql = Sql & " from facturas_variedad, usuarios.stipom stipom "
    Sql = Sql & " where facturas_variedad.numalbar = " & Albaran
    Sql = Sql & " and facturas_variedad.numlinealbar = " & Linea
    Sql = Sql & " and facturas_variedad.codtipom = stipom.codtipom "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Total = 0
    While Not Rs.EOF
        If vParamAplic.ContabilidadNueva Then
            Sql = "select sum(if(isnull(impvenci),0,impvenci) - if(isnull(impcobro),0,impcobro)) from cobros where numserie = " & DBSet(Rs.Fields(0).Value, "T")
            Sql = Sql & " and numfactu = " & DBSet(Rs.Fields(1).Value, "N")
            Sql = Sql & " and fecfactu = " & DBSet(Rs.Fields(2).Value, "F")
        Else
            Sql = "select sum(if(isnull(impvenci),0,impvenci) - if(isnull(impcobro),0,impcobro)) from scobro where numserie = " & DBSet(Rs.Fields(0).Value, "T")
            Sql = Sql & " and codfaccl = " & DBSet(Rs.Fields(1).Value, "N")
            Sql = Sql & " and fecfaccl = " & DBSet(Rs.Fields(2).Value, "F")
        End If
        Set Rs1 = New ADODB.Recordset
        Rs1.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs1.EOF And Not IsNull(Rs1.Fields(0)) Then
            Total = Total + DBLet(Rs1.Fields(0).Value, "N")
'++monica:10/02/2009 si me devuelve nulo no hay nada en la scobro
        Else
'            Exit Function
'++
        End If
    
        Rs.MoveNext
    Wend
    If Total = 0 Then
        FacturaSOCIO_CobradaTesoreria = 1
    End If
    Exit Function
    
eFacturaSOCIO_CobradaTesoreria:
    If Err.Number <> 0 Then
        FacturaSOCIO_CobradaTesoreria = 0
    End If
End Function


Public Function AlbaranSOCIO_Facturado(Albaran As Currency, Linea As Currency) As Byte
Dim Sql As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim Total As Currency

    On Error GoTo eAlbaranSOCIO_Facturado

    AlbaranSOCIO_Facturado = 0

    ' seleccionamos las facturas en donde aparece el albaran-linea
    Sql = "select count(*) "
    Sql = Sql & " from facturassocio_variedad "
    Sql = Sql & " where facturassocio_variedad.numalbar = " & Albaran
    Sql = Sql & " and facturassocio_variedad.numlinealbar = " & Linea
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0).Value, "N") > 0 Then AlbaranSOCIO_Facturado = 1
    End If
    
    Exit Function
eAlbaranSOCIO_Facturado:
    If Err.Number <> 0 Then
        AlbaranSOCIO_Facturado = 0
    End If
End Function

Public Function ImporteAlbaranSOCIO_Facturado(Albaran As Currency, Linea As Currency) As Double
Dim Sql As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim Total As Currency

    On Error GoTo eImporteAlbaranSOCIO_Facturado

    ImporteAlbaranSOCIO_Facturado = 0

    ' seleccionamos las facturas en donde aparece el albaran-linea
    Sql = "select sum(impornet) "
    Sql = Sql & " from facturassocio_variedad "
    Sql = Sql & " where facturassocio_variedad.numalbar = " & Albaran
    Sql = Sql & " and facturassocio_variedad.numlinealbar = " & Linea
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        ImporteAlbaranSOCIO_Facturado = DBLet(Rs.Fields(0).Value, "N")
    End If
    
    Exit Function
eImporteAlbaranSOCIO_Facturado:
    If Err.Number <> 0 Then
        ImporteAlbaranSOCIO_Facturado = 0
    End If
End Function

Public Function FacturasdeAlbaranSOCIO(Albaran As Currency, Linea As Currency) As String
Dim Sql As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim Total As Currency

    On Error GoTo eFacturasdeAlbaranSOCIO

    FacturasdeAlbaranSOCIO = ""

    ' seleccionamos las facturas en donde aparece el albaran-linea
    Sql = "select codtipom, numfactu, fecfactu "
    Sql = Sql & " from facturassocio_variedad "
    Sql = Sql & " where facturassocio_variedad.numalbar = " & Albaran
    Sql = Sql & " and facturassocio_variedad.numlinealbar = " & Linea
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
        Cad = Cad & "(" & DBSet(Rs.Fields(0).Value, "T") & "," & DBSet(Rs.Fields(1).Value, "N") & "," & DBSet(Rs.Fields(2).Value, "F") & "),"
        
        Rs.MoveNext
    Wend
    ' quitamos la ultima ,
    If Cad <> "" Then Cad = Mid(Cad, 1, Len(Cad) - 1)
    
    '[Monica]05/03/2013: indicamos la tabla pq sino da error en el vista previa
    FacturasdeAlbaranSOCIO = "(facturassocio.codtipom, facturassocio.numfactu, facturassocio.fecfactu) in (" & Cad & ")"
    Exit Function
    
eFacturasdeAlbaranSOCIO:
    If Err.Number <> 0 Then
        FacturasdeAlbaranSOCIO = ""
    End If
End Function



Public Function ImporteAlbaranFacturadoNoCobrado(Albaran As Currency, Linea As Currency, Parcial As Boolean) As Double
Dim Sql As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim Total As Currency
Dim ImporteCobrado As Currency
Dim Cad As String

    On Error GoTo eImporteAlbaranFacturadoNoCobrado

    ImporteAlbaranFacturadoNoCobrado = 0

    ' seleccionamos las facturas en donde aparece el albaran-linea
    Sql = "select sum(impornet) "
    Sql = Sql & " from facturas_variedad "
    Sql = Sql & " where facturas_variedad.numalbar = " & Albaran
    Sql = Sql & " and facturas_variedad.numlinealbar = " & Linea
    
    Cad = FacturasCobradasEnTesoreria(Albaran, Linea, Total, ImporteCobrado)
    '[Monica]10/04/2012: he añadido la condicion de factura contabilizada
    If Cad <> "" And FacturaContabilizada(Albaran, Linea) Then
        Sql = Sql & " and (codtipom, numfactu, fecfactu) not in (" & Cad & ")"
    End If
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        If Total = 0 Or ImporteCobrado = 0 Then
            ImporteAlbaranFacturadoNoCobrado = DBLet(Rs.Fields(0).Value, "N")
            Parcial = False
        Else
            ImporteAlbaranFacturadoNoCobrado = Total
            Parcial = True
        End If
    End If
    
    Exit Function
eImporteAlbaranFacturadoNoCobrado:
    If Err.Number <> 0 Then
        ImporteAlbaranFacturadoNoCobrado = 0
    End If
End Function


Public Function FacturasCobradasEnTesoreria(Albaran As Currency, Linea As Currency, Importe As Currency, ImporteCobrado As Currency) As String
Dim Sql As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim CADENA As String
Dim Albaranes As Long
    
    On Error GoTo eFacturasCobradasEnTesoreria

    FacturasCobradasEnTesoreria = ""

    ' seleccionamos las facturas en donde aparece el albaran-linea
    Sql = "select distinct stipom.letraser, facturas_variedad.numfactu, facturas_variedad.fecfactu, facturas_variedad.codtipom "
    Sql = Sql & " from facturas_variedad, usuarios.stipom stipom "
    Sql = Sql & " where facturas_variedad.numalbar = " & Albaran
    Sql = Sql & " and facturas_variedad.numlinealbar = " & Linea
    Sql = Sql & " and facturas_variedad.codtipom = stipom.codtipom "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CADENA = ""
    Importe = 0
    ImporteCobrado = 0
    Albaranes = 0
    While Not Rs.EOF
        If vParamAplic.ContabilidadNueva Then
            Sql = "select sum(if(isnull(impvenci),0,impvenci) - if(isnull(impcobro),0,impcobro)), sum(if(isnull(impcobro),0,impcobro)) from cobros where numserie = " & DBSet(Rs.Fields(0).Value, "T")
            Sql = Sql & " and numfactu = " & DBSet(Rs.Fields(1).Value, "N")
            Sql = Sql & " and fecfactu = " & DBSet(Rs.Fields(2).Value, "F")
        Else
            Sql = "select sum(if(isnull(impvenci),0,impvenci) - if(isnull(impcobro),0,impcobro)), sum(if(isnull(impcobro),0,impcobro)) from scobro where numserie = " & DBSet(Rs.Fields(0).Value, "T")
            Sql = Sql & " and codfaccl = " & DBSet(Rs.Fields(1).Value, "N")
            Sql = Sql & " and fecfaccl = " & DBSet(Rs.Fields(2).Value, "F")
        End If
        Set Rs1 = New ADODB.Recordset
        Rs1.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        
'        ' cuantos albaranes hay colgando de esta factura
'        Sql = "select count(*) from facturas_variedad where facturas_variedad.codtipom = " & DBSet(RS.Fields(3).Value, "T") & " and facturas_variedad.numfactu = " & DBSet(RS.Fields(1).Value, "N") & " and fecfactu = " & DBSet(RS.Fields(2).Value, "F")
'        Albaranes = DevuelveValor(Sql)
        
        If Not Rs1.EOF And Not IsNull(Rs1.Fields(0)) Then
            Importe = Importe + DBLet(Rs1.Fields(0).Value, "N")
            ImporteCobrado = ImporteCobrado + DBLet(Rs1.Fields(1).Value, "N")
        Else
            CADENA = CADENA & "(" & DBSet(Rs.Fields(3).Value, "T") & "," & DBSet(Rs.Fields(1).Value, "N") & "," & DBSet(Rs.Fields(2).Value, "F") & "),"
        End If
    
        Rs.MoveNext
    Wend
    If CADENA <> "" Then
        CADENA = Mid(CADENA, 1, Len(CADENA) - 1)
        FacturasCobradasEnTesoreria = CADENA
    End If
    Exit Function
    
eFacturasCobradasEnTesoreria:
    If Err.Number <> 0 Then
        FacturasCobradasEnTesoreria = ""
    End If
End Function






