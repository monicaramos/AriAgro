Attribute VB_Name = "ModContabilizar"
Option Explicit


'===================================================================================
'CONTABILIZAR FACTURAS:
'Modulo para el traspaso de registros de cabecera y lineas de tablas de FACTURACION
'A las tablas de FACTURACION de Contabilidad
'====================================================================================

Private DtoGnral As Currency
Private DtoPPago As Currency
Private BaseImp As Currency
Private IvaImp As Currency
Private TotalFac As Currency
Private CCoste As String
Private conCtaAlt As Boolean 'el cliente utiliza cuentas alternativas

'Para pasar a contabilidad facturas de proveedor
Private AnyoFacPr As Integer 'a�o factura proveedor, es el ano de fecha_recepcion

Dim vvIban As String

Private vTipoIva(2) As Currency
Private vPorcIva(2) As Currency
Private vPorcRec(2) As Currency
Private vBaseIva(2) As Currency
Private vImpIva(2) As Currency
Private vImpRec(2) As Currency

Private vvFecRecep As String

Public Function CrearTMPFacturas(cadTABLA As String, cadWHERE As String) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim sql As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturas = False
    
    sql = "CREATE TEMPORARY TABLE tmpFactu ( "
    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
        sql = sql & "codtipom char(3) NOT NULL default '',"
        sql = sql & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        If cadTABLA = "scafpc" Then
            sql = sql & "codprove int(6) unsigned NOT NULL default '0',"
            sql = sql & "numfactu varchar(10)  NOT NULL ,"
        Else
            sql = sql & "codtrans smallint(3) unsigned NOT NULL default '0',"
            sql = sql & "numfactu varchar(10)  NOT NULL ,"
        End If
    End If
    sql = sql & "fecfactu date NOT NULL default '0000-00-00') "
    conn.Execute sql
     
     
    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
        sql = "SELECT codtipom, numfactu, fecfactu"
    Else
        If cadTABLA = "scafpc" Then
            sql = "SELECT codprove, numfactu, fecfactu"
        Else
            sql = "SELECT codtrans, numfactu, fecfactu"
        End If
    End If
    sql = sql & " FROM " & cadTABLA
    sql = sql & " WHERE " & cadWHERE
    sql = " INSERT INTO tmpFactu " & sql
    conn.Execute sql

    CrearTMPFacturas = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturas = False
        'Borrar la tabla temporal
        sql = " DROP TABLE IF EXISTS tmpFactu;"
        conn.Execute sql
    End If
End Function


Public Sub BorrarTMPFacturas()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpFactu;"
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub InsertarTMPErrFac(MenError As String, cadWHERE As String)
Dim sql As String

    On Error Resume Next
    sql = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
    sql = sql & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    sql = sql & " WHERE " & Replace(cadWHERE, "scafpc", "tmpFactu")
    conn.Execute sql
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Function CrearTMPErrFact(cadTABLA As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim sql As String
    
    On Error GoTo ECrear
    
    CrearTMPErrFact = False
    
    sql = "CREATE TEMPORARY TABLE tmpErrFac ( "
    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
        sql = sql & "codtipom char(3) NOT NULL default '',"
        sql = sql & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        sql = sql & "codprove int(6) unsigned NOT NULL default '0',"
        sql = sql & "numfactu varchar(10) NOT NULL ,"
    End If
    sql = sql & "fecfactu date NOT NULL default '0000-00-00', "
    sql = sql & "error varchar(200) NULL )"
    conn.Execute sql
     
     CrearTMPErrFact = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrFact = False
        'Borrar la tabla temporal
        sql = " DROP TABLE IF EXISTS tmpErrFac;"
        conn.Execute sql
    End If
End Function


Public Sub BorrarTMPErrFact()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmpErrFac;"
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function ComprobarLetraSerie(cadTABLA As String) As Boolean
'Para Facturas VENTA a clientes
'Comprueba que la letra del serie del tipo de movimiento es  correcta
Dim sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cad As String, devuelve As String

On Error GoTo EComprobarLetra

    ComprobarLetraSerie = False
    
    'Comprobar que existe la letra de serie en contabilidad
    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
        'cargamos el RSConta con la tabla contadores de BD: Contabilidad
        'donde estan todas las letra de serie que existen en la contabilidad
        sql = "Select distinct tiporegi from contadores"
        Set RSconta = New ADODB.Recordset
        RSconta.Open sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
        If RSconta.EOF Then
            RSconta.Close
            Set RSconta = Nothing
            Exit Function
        End If
            
    
        'obtenemos los distintos tipos de movimiento que vamos a contabilizar
        'de las facturas seleccionadas
        sql = "select distinct " & cadTABLA & ".codtipom from " & cadTABLA
        sql = sql & " INNER JOIN tmpFactu ON " & cadTABLA & ".codtipom=tmpFactu.codtipom AND " & cadTABLA & ".numfactu=tmpFactu.numfactu AND " & cadTABLA & ".fecfactu=tmpFactu.fecfactu "
'        SQL = SQL & cadWHERE
        Set Rs = New ADODB.Recordset
        Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cad = ""
        b = True
        While Not Rs.EOF And b
            'comprobar que todas las letras serie existen en Ariges
'--monica:10/02/2009
'            SQL = "letraser"
'            devuelve = DevuelveDesdeBDNew(cAgro, "stipom", "codtipom", "codtipom", Rs!codTipoM, "T", SQL)
'++monica:10/02/2009
            sql = ObtenerLetraSerie(Rs!codTipoM)
            devuelve = DBLet(Rs!codTipoM, "T")
'++
            If devuelve = "" Then
                b = False
                cad = Rs!codTipoM & " en BD de Ariagro."
            ElseIf sql <> "" Then
                'comprobar que todas las letras serie existen en la contabilidad
                devuelve = "tiporegi= " & DBSet(sql, "T")
                RSconta.MoveFirst
                RSconta.Find (devuelve), , adSearchForward
                If RSconta.EOF Then
                    'no encontrado
                    b = False
                    cad = sql & " en BD de Contabilidad."
                End If
            End If
            If b Then cad = cad & DBSet(Rs!codTipoM, "T") & ","
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        RSconta.Close
        Set RSconta = Nothing
        
        If Not b Then 'Hay algun movimiento que no existe
            devuelve = "No existe el tipo de movimiento: " & cad & vbCrLf
            devuelve = devuelve & "Consulte con el administrador."
            MsgBox devuelve, vbExclamation
            Exit Function
        End If
        
        'Todos los Tipo de movimiento existen
        If cad <> "" Then
            cad = Mid(cad, 1, Len(cad) - 1) 'quitamos ult. coma
        
            'miramos si hay algun movimiento de factura que la letra serie sea nulo
            sql = "select count(*) from usuarios.stipom "
            sql = sql & "where codtipom IN (" & cad & ") and (isnull(letraser) or letraser='')"
            If RegistrosAListar(sql) > 0 Then
                sql = "Hay algun tipo de movimiento de Facturaci�n que no tiene letra serie." & vbCrLf
                sql = sql & "Comprobar en la tabla de tipos de movimiento: " & cad
                MsgBox sql, vbExclamation
                Exit Function
            End If
        End If
        ComprobarLetraSerie = True
    Else
        ComprobarLetraSerie = True
    End If

EComprobarLetra:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Letra Serie", Err.Description
    End If
End Function

'###### ESTE YA NO SE UTILIZA
'Public Function ComprobarNumFacturas(cadTabla As String, cadWConta) As Boolean
''Comprobar que no exista ya en la contabilidad un n� de factura para la fecha que
''vamos a contabilizar
'Dim SQL As String
'Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
'Dim b As Boolean
'
'    On Error GoTo ECompFactu
'
'    ComprobarNumFacturas = False
'
'    SQL = "SELECT numserie,codfaccl,anofaccl FROM cabfact "
'    SQL = SQL & " WHERE " & cadWConta
'
'    Set RSconta = New ADODB.Recordset
'    RSconta.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    If Not RSconta.EOF Then
'        'Seleccionamos las distintas facturas que vamos a facturar
'        SQL = "SELECT DISTINCT " & cadTabla & ".codtipom,letraser,scafac.numfactu,scafac.fecfactu "
'        SQL = SQL & " FROM (" & cadTabla & " INNER JOIN stipom ON " & cadTabla & ".codtipom=stipom.codtipom) "
'        SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
''        SQL = SQL & " WHERE " & cadWHERE
'
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        b = True
'        While Not RS.EOF And b
'            SQL = "(numserie= " & DBSet(RS!LetraSer, "T") & " AND codfaccl=" & DBSet(RS!NumFactu, "N") & " AND anofaccl=" & Year(RS!FecFactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
'                b = False
'                SQL = "          N� Fac.: " & Format(RS!NumFactu, "0000000") & vbCrLf
'                SQL = SQL & "          Fecha: " & RS!FecFactu
'            End If
'            RS.MoveNext
'        Wend
'        RS.Close
'        Set RS = Nothing
'
'        If Not b Then
'            SQL = "Ya existe la factura: " & vbCrLf & SQL
'            SQL = "Comprobando N� Facturas en Contabilidad...       " & vbCrLf & vbCrLf & SQL
'
'            MsgBox SQL, vbExclamation
'            ComprobarNumFacturas = False
'        Else
'            ComprobarNumFacturas = True
'        End If
'    Else
'        ComprobarNumFacturas = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
'
'ECompFactu:
'     If Err.Number <> 0 Then
'        MuestraError Err.Number, "Comprobar N� Facturas", Err.Description
'    End If
'End Function


Public Function ComprobarNumFacturas_new(cadTABLA As String, cadWConta) As Boolean
'Comprobar que no exista ya en la contabilidad un n� de factura para la fecha que
'vamos a contabilizar
Dim sql As String
Dim SQLconta As String
Dim Rs As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECompFactu

    ComprobarNumFacturas_new = False
    
'    SQLconta = "SELECT numserie,codfaccl,anofaccl FROM cabfact "
    If vParamAplic.ContabilidadNueva Then
        SQLconta = "SELECT count(*) FROM factcli WHERE "
    Else
        SQLconta = "SELECT count(*) FROM cabfact WHERE "
    End If
'    SQLconta = SQLconta & " WHERE (" & cadWConta & ") "

    
'    Set RSconta = New ADODB.Recordset
'    RSconta.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText

'    If Not RSconta.EOF Then
        'Seleccionamos las distintas facturas que vamos a facturar
        sql = "SELECT DISTINCT " & cadTABLA & ".codtipom,letraser," & cadTABLA & ".numfactu," & cadTABLA & ".fecfactu "
        sql = sql & " FROM (" & cadTABLA & " INNER JOIN usuarios.stipom stipom ON " & cadTABLA & ".codtipom=stipom.codtipom) "
        sql = sql & " INNER JOIN tmpFactu ON " & cadTABLA & ".codtipom=tmpFactu.codtipom AND " & cadTABLA & ".numfactu=tmpFactu.numfactu AND " & cadTABLA & ".fecfactu=tmpFactu.fecfactu "

        
        Set Rs = New ADODB.Recordset
        Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF And b
            If vParamAplic.ContabilidadNueva Then
                sql = "(numserie= " & DBSet(Rs!letraser, "T") & " AND numfactu=" & DBSet(Rs!NumFactu, "N") & " AND anofactu=" & Year(Rs!FecFactu) & ")"
            Else
                sql = "(numserie= " & DBSet(Rs!letraser, "T") & " AND codfaccl=" & DBSet(Rs!NumFactu, "N") & " AND anofaccl=" & Year(Rs!FecFactu) & ")"
            End If
'            If SituarRSetMULTI(RSconta, SQL) Then
            sql = SQLconta & sql
            If RegistrosAListar(sql, cConta) Then
                b = False
                sql = "          Letra Serie: " & DBSet(Rs!letraser, "T") & vbCrLf
                sql = sql & "          N� Fac.: " & Format(Rs!NumFactu, "0000000") & vbCrLf
                sql = sql & "          Fecha: " & Rs!FecFactu
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            sql = "Ya existe la factura: " & vbCrLf & sql
            sql = "Comprobando N� Facturas en Contabilidad...       " & vbCrLf & vbCrLf & sql
            
            MsgBox sql, vbExclamation
            ComprobarNumFacturas_new = False
        Else
            ComprobarNumFacturas_new = True
        End If
'    Else
'        ComprobarNumFacturas_new = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
    Exit Function
    
ECompFactu:
     If Err.Number <> 0 Then
        ComprobarNumFacturas_new = False
        MuestraError Err.Number, "Comprobar N� Facturas", Err.Description
    End If
End Function




'###### ESTE YA NO SE UTILIZA
'Public Function ComprobarCtaContable(cadTabla As String, Opcion As Byte) As Boolean
''Comprobar que todas las ctas contables de los distintos clientes de las facturas
''que vamos a contabilizar existan en la contabilidad
'Dim SQL As String
'Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
'Dim b As Boolean
'Dim cadG As String
'
'    On Error GoTo ECompCta
'
'    ComprobarCtaContable = False
'
'    If Opcion = 3 Then 'si hay analitica comprobar que todas las cuentas
'                        'empiezan por el digito que hay en conta.parametros.grupogto o .grupovta
'        cadG = "grupovta"
'        SQL = DevuelveDesdeBDNew(conConta, "parametros", "grupogto", "", "", "", cadG)
'        If SQL <> "" And cadG <> "" Then
'            SQL = " AND (codmacta like '" & SQL & "%' OR codmacta like '" & cadG & "%')"
'        ElseIf SQL <> "" Then
'            SQL = " AND (codmacta like '" & SQL & "%')"
'        ElseIf cadG <> "" Then
'            SQL = " AND (codmacta like '" & cadG & "%')"
'        End If
'        cadG = SQL
'    End If
'
'    SQL = "SELECT codmacta FROM cuentas "
'    SQL = SQL & " WHERE apudirec='S'"
'    If cadG <> "" Then SQL = SQL & cadG
'
'    Set RSconta = New ADODB.Recordset
'    RSconta.Open SQL, ConnConta, adOpenStatic, adLockPessimistic, adCmdText
'
'    If Not RSconta.EOF Then
'        If Opcion = 1 Then
'            If cadTabla = "scafac" Then
'                'Seleccionamos los distintos clientes,cuentas que vamos a facturar
'                SQL = "SELECT DISTINCT scafac.codclien, sclien.codmacta "
'                SQL = SQL & " FROM (scafac INNER JOIN sclien ON scafac.codclien=sclien.codclien) "
'                SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
'            Else
'                'Seleccionamos los distintos proveedores,cuentas que vamos a facturar
'                SQL = "SELECT DISTINCT scafpc.codprove, sprove.codmacta "
'                SQL = SQL & " FROM (scafpc INNER JOIN sprove ON scafpc.codprove=sprove.codprove) "
'                SQL = SQL & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
'            End If
'
'        ElseIf Opcion = 2 Or Opcion = 3 Then
'            SQL = "SELECT distinct "
'            If Opcion = 2 Then SQL = SQL & " sartic.codfamia,"
'            If cadTabla = "scafac" Then
'                SQL = SQL & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((slifac "
'                SQL = SQL & " INNER JOIN tmpFactu ON slifac.codtipom=tmpFactu.codtipom AND slifac.numfactu=tmpFactu.numfactu AND slifac.fecfactu=tmpFactu.fecfactu) "
'                SQL = SQL & "INNER JOIN sartic ON slifac.codartic=sartic.codartic) "
'            Else
'                SQL = SQL & " sfamia.ctacompr as codmacta,sfamia.abocompr as ctaabono from ((slifpc "
'                SQL = SQL & " INNER JOIN tmpFactu ON slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu) "
'                SQL = SQL & "INNER JOIN sartic ON slifpc.codartic=sartic.codartic) "
'            End If
'            SQL = SQL & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
'        End If
'
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        b = True
'        While Not RS.EOF And b
'            SQL = "codmacta= " & DBSet(RS!Codmacta, "T")
'            RSconta.MoveFirst
'            RSconta.Find (SQL), , adSearchForward
'            If RSconta.EOF Then
'                b = False 'no encontrado
'                If Opcion = 1 Then
'                    If cadTabla = "scafac" Then
'                        SQL = RS!Codmacta & " del Cliente " & Format(RS!CodClien, "000000")
'                    Else
'                        SQL = RS!Codmacta & " del Proveedor " & Format(RS!codProve, "000000")
'                    End If
'                ElseIf Opcion = 2 Then
'                    SQL = RS!Codmacta & " de la familia " & Format(RS!codfamia, "0000")
'                ElseIf Opcion = 3 Then
'                    SQL = RS!Codmacta
'                End If
'            End If
'
'            If Opcion = 2 Then
'                'Comprobar que ademas de existir la cuenta de ventas exista tambien
'                'la cuenta ABONO ventas
'                SQL = "codmacta= " & DBSet(RS!ctaabono, "T")
'                RSconta.MoveFirst
'                RSconta.Find (SQL), , adSearchForward
'                If RSconta.EOF Then
'                    b = False 'no encontrado
'                    SQL = RS!ctaabono & " de la familia " & Format(RS!codfamia, "0000")
'                End If
'            End If
'
'            'comprobar cuentas alternativas solo para facturacion a clientes
'            If cadTabla = "scafac" Then
'                If Opcion = 2 Then
'                    ' Comprobar cuenta venta alternativa
'                    If DBLet(RS!ctavent1, "T") <> "" Then
'                        SQL = "codmacta= " & DBSet(RS!ctavent1, "T")
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
'                            b = False 'no encontrado
'                            SQL = RS!ctavent1 & " de la familia " & Format(RS!codfamia, "0000")
'                        End If
'                    Else
'                        b = False
'                        SQL = " o la familia no tiene asignada cuenta venta alternativa."
'                    End If
'                End If
'                If Opcion = 2 Then
'                    ' Comprobar cuenta de abono alternativa
'                    If DBLet(RS!abovent1, "T") <> "" Then
'                        SQL = "codmacta= " & DBSet(RS!abovent1, "T")
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
'                            b = False 'no encontrado
'                            SQL = RS!ctaabon1 & " de la familia " & Format(RS!codfamia, "0000")
'                        End If
'                    Else
'                        b = False
'                        SQL = " o la familia no tiene asignada cuenta abono alternativa."
'                    End If
'                End If
'            End If
'            RS.MoveNext
'        Wend
'        RS.Close
'        Set RS = Nothing
'
'        If Not b Then
'            If Opcion <> 3 Then
'                SQL = "No existe la cta contable " & SQL
'            Else
'                SQL = "La cuenta " & SQL & " no es del nivel correcto."
'            End If
'            SQL = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & SQL
'
'            MsgBox SQL, vbExclamation
'            ComprobarCtaContable = False
'        Else
'            ComprobarCtaContable = True
'        End If
'    Else
'        ComprobarCtaContable = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
'
'ECompCta:
'     If Err.Number <> 0 Then
'        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
'    End If
'End Function






Public Function ComprobarCtaContable_new(cadTABLA As String, Opcion As Byte) As Boolean
'Comprobar que todas las ctas contables de los distintos clientes de las facturas
'que vamos a contabilizar existan en la contabilidad
Dim sql As String
Dim Rs As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cadG As String
Dim SQLcuentas As String
Dim CadCampo1 As String
Dim numNivel As String
Dim NumDigit As String
Dim NumDigit3 As String

Dim SeccionHorto As Integer

    On Error GoTo ECompCta

    ComprobarCtaContable_new = False
    
    cadG = ""
    If Opcion = 3 Or Opcion = 7 Or Opcion = 10 Or Opcion = 13 Then
        'si hay analitica comprobar que todas las cuentas
        'empiezan por el digito que hay en conta.parametros.grupogto o .grupovta
        cadG = "grupovta"
        sql = DevuelveDesdeBDNew(cConta, "parametros", "grupogto", "", "", "", cadG)
        If sql <> "" And cadG <> "" Then
            sql = " AND (codmacta like '" & sql & "%' OR codmacta like '" & cadG & "%')"
        ElseIf sql <> "" Then
            sql = " AND (codmacta like '" & sql & "%')"
        ElseIf cadG <> "" Then
            sql = " AND (codmacta like '" & cadG & "%')"
        End If
        cadG = sql
    End If
    
    
'    SQL = "SELECT codmacta FROM cuentas "
'    SQL = SQL & " WHERE apudirec='S'"
'    If cadG <> "" Then SQL = SQL & cadG
    
    SQLcuentas = "SELECT count(*) FROM cuentas WHERE apudirec='S' "
    If cadG <> "" Then SQLcuentas = SQLcuentas & cadG
    
    If Opcion = 1 Then
        If cadTABLA = "facturas" Then
            'Seleccionamos los distintos clientes,cuentas que vamos a facturar
            sql = "SELECT DISTINCT facturas.codclien, clientes.codmacta "
            sql = sql & " FROM (facturas INNER JOIN clientes ON facturas.codclien=clientes.codclien) "
            sql = sql & " INNER JOIN tmpFactu ON facturas.codtipom=tmpFactu.codtipom AND facturas.numfactu=tmpFactu.numfactu AND facturas.fecfactu=tmpFactu.fecfactu "
        Else
            If cadTABLA = "facturassocio" Then
                SeccionHorto = DevuelveValor("select seccionhorto from rparam")
                'Seleccionamos las distintas cuentas de clientes de la seccion de horto, de los socios que vamos a facturar
                sql = "SELECT DISTINCT facturassocio.codsocio, rsocios_seccion.codmaccli codmacta "
                sql = sql & " FROM (facturassocio INNER JOIN rsocios_seccion ON facturassocio.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & SeccionHorto & ") "
                sql = sql & " INNER JOIN tmpFactu ON facturassocio.codtipom=tmpFactu.codtipom AND facturassocio.numfactu=tmpFactu.numfactu AND facturassocio.fecfactu=tmpFactu.fecfactu "
            Else
                If cadTABLA = "scafpc" Then
                    'Seleccionamos los distintos proveedores,cuentas que vamos a facturar
                    sql = "SELECT DISTINCT scafpc.codprove, proveedor.codmacta "
                    sql = sql & " FROM (scafpc INNER JOIN proveedor ON scafpc.codprove=proveedor.codprove) "
                    sql = sql & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
                Else
                    'Seleccionamos los distintos transportistas ,cuentas que vamos a facturar
                    sql = "SELECT DISTINCT tcafpc.codtrans, agencias.codmacta "
                    sql = sql & " FROM (tcafpc INNER JOIN agencias ON tcafpc.codtrans=agencias.codtrans) "
                    sql = sql & " INNER JOIN tmpFactu ON tcafpc.codtrans=tmpFactu.codtrans AND tcafpc.numfactu=tmpFactu.numfactu AND tcafpc.fecfactu=tmpFactu.fecfactu "
                
                End If
            End If
        End If
    ElseIf Opcion = 2 Or Opcion = 3 Or Opcion = 8 Then
        sql = "SELECT distinct "
        If Opcion = 2 Then sql = sql & " sartic.codfamia,"
        If cadTABLA = "facturas" Then
            If Opcion <> 8 Then
                sql = sql & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((facturas_envases "
                sql = sql & " INNER JOIN tmpFactu ON facturas_envases.codtipom=tmpFactu.codtipom AND facturas_envases.numfactu=tmpFactu.numfactu AND facturas_envases.fecfactu=tmpFactu.fecfactu) "
                sql = sql & "INNER JOIN sartic ON facturas_envases.codartic=sartic.codartic) "
            Else
                numNivel = DevuelveDesdeBDNew(cConta, "empresa", "numnivel", "codempre", vParamAplic.NumeroConta, "N")
                NumDigit = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & numNivel, "codempre", vParamAplic.NumeroConta, "N")
                NumDigit3 = DevuelveDesdeBDNew(cConta, "empresa", "numdigi3", "codempre", vParamAplic.NumeroConta, "N")
                
'                CadCampo1 = "concat(concat(variedades.raizctavtas,tipomer.digicont), right(concat('0000000000',albaran_variedad.codvarie)," & (CCur(NumDigit) - CCur(NumDigit3) - 1) & "))"
                CadCampo1 = "CASE tipomer.tiptimer WHEN 0 THEN ctavtasinterior WHEN 1 THEN ctavtasexportacion WHEN 2 THEN ctavtasindustria WHEN 3 THEN ctavtasretirada WHEN 4 THEN ctavtasotros END"
                
                sql = sql & " albaran_variedad.codvarie, " & CadCampo1 & " as codmacta from ((((((facturas_variedad "
                sql = sql & " INNER JOIN tmpFactu ON facturas_variedad.codtipom=tmpFactu.codtipom AND facturas_variedad.numfactu=tmpFactu.numfactu AND facturas_variedad.fecfactu=tmpFactu.fecfactu) "
                sql = sql & " inner join usuarios.stipom stipom on facturas_variedad.codtipom=stipom.codtipom) "
                sql = sql & " inner join albaran on facturas_variedad.numalbar = albaran.numalbar) "
                sql = sql & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
                sql = sql & " inner join albaran_variedad on facturas_variedad.numalbar = albaran_variedad.numalbar and facturas_variedad.numlinealbar = albaran_variedad.numlinea) "
                sql = sql & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
                
                
'                Sql = Sql & " INNER JOIN tmpFactu ON facturas_variedad.codtipom=tmpFactu.codtipom AND facturas_variedad.numfactu=tmpFactu.numfactu AND facturas_variedad.fecfactu=tmpFactu.fecfactu) "
'                Sql = Sql & "INNER JOIN sartic ON facturas_envases.codartic=sartic.codartic) "
            End If
        Else
            If cadTABLA = "facturassocio" Then
                If Opcion <> 8 Then
                    sql = sql & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((facturassocio_envases "
                    sql = sql & " INNER JOIN tmpFactu ON facturassocio_envases.codtipom=tmpFactu.codtipom AND facturassocio_envases.numfactu=tmpFactu.numfactu AND facturassocio_envases.fecfactu=tmpFactu.fecfactu) "
                    sql = sql & "INNER JOIN sartic ON facturassocio_envases.codartic=sartic.codartic) "
                Else
                    numNivel = DevuelveDesdeBDNew(cConta, "empresa", "numnivel", "codempre", vParamAplic.NumeroConta, "N")
                    NumDigit = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & numNivel, "codempre", vParamAplic.NumeroConta, "N")
                    NumDigit3 = DevuelveDesdeBDNew(cConta, "empresa", "numdigi3", "codempre", vParamAplic.NumeroConta, "N")
                    
    '                CadCampo1 = "concat(concat(variedades.raizctavtas,tipomer.digicont), right(concat('0000000000',albaran_variedad.codvarie)," & (CCur(NumDigit) - CCur(NumDigit3) - 1) & "))"
                    CadCampo1 = "CASE tipomer.tiptimer WHEN 0 THEN ctavtasinterior WHEN 1 THEN ctavtasexportacion WHEN 2 THEN ctavtasindustria WHEN 3 THEN ctavtasretirada WHEN 4 THEN ctavtasotros END"
                    
                    sql = sql & " albaran_variedad.codvarie, " & CadCampo1 & " as codmacta from ((((((facturassocio_variedad "
                    sql = sql & " INNER JOIN tmpFactu ON facturassocio_variedad.codtipom=tmpFactu.codtipom AND facturassocio_variedad.numfactu=tmpFactu.numfactu AND facturassocio_variedad.fecfactu=tmpFactu.fecfactu) "
                    sql = sql & " inner join usuarios.stipom stipom on facturassocio_variedad.codtipom=stipom.codtipom) "
                    sql = sql & " inner join albaran on facturassocio_variedad.numalbar = albaran.numalbar) "
                    sql = sql & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
                    sql = sql & " inner join albaran_variedad on facturassocio_variedad.numalbar = albaran_variedad.numalbar and facturassocio_variedad.numlinealbar = albaran_variedad.numlinea) "
                    sql = sql & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
                End If
            
            Else
                sql = sql & " sfamia.ctacompr as codmacta,sfamia.abocompr as ctaabono from ((slifpc "
                sql = sql & " INNER JOIN tmpFactu ON slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu) "
                sql = sql & "INNER JOIN sartic ON slifpc.codartic=sartic.codartic) "
            End If
        End If
        If Opcion <> 8 Then sql = sql & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
    ElseIf Opcion = 4 Or Opcion = 6 Then
        sql = "select distinct " & DBSet(vParamAplic.CtaTraReten, "T") & " as codmacta from tcafpc "
    ElseIf Opcion = 5 Or Opcion = 7 Then
'        Sql = "select distinct " & DBSet(vParamAplic.CtaAboTrans, "T") & " as codmacta from tcafpc "
'       transporte
            sql = " SELECT if(tipomer.tiptimer = 1,variedades.ctatraexporta,variedades.ctatrainterior) as cuenta "
            sql = sql & " FROM tlifpc, albaran, albaran_variedad, variedades, tipomer, tmpFactu, tcafpc  WHERE "
            sql = sql & " tcafpc.tipo = 0 and " ' transportista
            sql = sql & " tlifpc.codtrans=tmpFactu.codtrans and tlifpc.numfactu=tmpFactu.numfactu and tlifpc.fecfactu=tmpFactu.fecfactu and "
            sql = sql & " tlifpc.numalbar=albaran_variedad.numalbar and "
            sql = sql & " tlifpc.numlinea=albaran_variedad.numlinea and "
            sql = sql & " tlifpc.codtrans=tcafpc.codtrans and tlifpc.numfactu=tcafpc.numfactu and tlifpc.fecfactu=tcafpc.fecfactu and "
            sql = sql & " albaran_variedad.numalbar=albaran.numalbar and "
            sql = sql & " albaran_variedad.codvarie=variedades.codvarie and "
            sql = sql & " albaran.codtimer=tipomer.codtimer "
            sql = sql & " group by 1 "

    ElseIf Opcion = 12 Or Opcion = 13 Then
'       comisionista
            sql = " SELECT variedades.ctacomisionista as cuenta, variedades.codvarie  "
            sql = sql & " FROM tlifpc, albaran, albaran_variedad, variedades, tipomer, tmpFactu, tcafpc  WHERE "
            sql = sql & " tcafpc.tipo = 1 and " ' comisionista
            sql = sql & " tlifpc.codtrans=tmpFactu.codtrans and tlifpc.numfactu=tmpFactu.numfactu and tlifpc.fecfactu=tmpFactu.fecfactu and "
            sql = sql & " tlifpc.numalbar=albaran_variedad.numalbar and "
            sql = sql & " tlifpc.numlinea=albaran_variedad.numlinea and "
            sql = sql & " tlifpc.codtrans=tcafpc.codtrans and tlifpc.numfactu=tcafpc.numfactu and tlifpc.fecfactu=tcafpc.fecfactu and "
            sql = sql & " albaran_variedad.numalbar=albaran.numalbar and "
            sql = sql & " albaran_variedad.codvarie=variedades.codvarie and "
            sql = sql & " albaran.codtimer=tipomer.codtimer "
            sql = sql & " group by 1 "
            
    ElseIf Opcion = 9 Or Opcion = 10 Then
            sql = " select codmacta as cuenta "
            sql = sql & " from tcafpv, tmpFactu "
            sql = sql & " where tmpFactu.codtrans=tcafpv.codtrans and tmpFactu.numfactu=tcafpv.numfactu and tmpFactu.fecfactu=tcafpv.fecfactu "
            sql = sql & " group by 1 "
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    sql = ""
    b = True

    While Not Rs.EOF And b
        If Opcion < 4 Or Opcion = 8 Then
            sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!Codmacta, "T")
        ElseIf Opcion = 4 Or Opcion = 6 Then
            sql = SQLcuentas & " AND codmacta= " & DBSet(vParamAplic.CtaTraReten, "T")
        ElseIf Opcion = 5 Or Opcion = 7 Then
            sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!Cuenta, "T")
        ElseIf Opcion = 12 Or Opcion = 13 Then
            sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!Cuenta, "T")
        ElseIf Opcion = 9 Or Opcion = 10 Then
            sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!Cuenta, "T")
        End If
            
        
        If Not (RegistrosAListar(sql, cConta) > 0) Then
        'si no lo encuentra
            b = False 'no encontrado
            If Opcion = 1 Then
                If cadTABLA = "facturas" Then
                    sql = Rs!Codmacta & " del Cliente " & Format(Rs!CodClien, "000000")
                Else
                    If cadTABLA = "facturassocio" Then
                        sql = Rs!Codmacta & " del Socio " & Format(Rs!CodSocio, "000000")
                    Else
                        If cadTABLA = "scafpc" Then
                            sql = Rs!Codmacta & " del Proveedor " & Format(Rs!codProve, "000000")
                        Else
                            sql = Rs!Codmacta & " del Transportista " & Format(Rs!codTrans, "000")
                        End If
                    End If
                End If
            ElseIf Opcion = 2 Then
                sql = Rs!Codmacta & " de la familia " & Format(Rs!codfamia, "0000")
            ElseIf Opcion = 3 Then
                sql = Rs!Codmacta
            ElseIf Opcion = 4 Or Opcion = 6 Then
                sql = vParamAplic.CtaTraReten
            ElseIf Opcion = 5 Or Opcion = 7 Then
                sql = DBLet(Rs!Cuenta, "T") ' vParamAplic.CtaAboTrans
            ElseIf Opcion = 12 Or Opcion = 13 Then
                sql = DBLet(Rs!Cuenta, "T") & " de comisionista de la variedad " & Format(Rs!codvarie, "000000")
            ElseIf Opcion = 8 Then
                sql = Rs!Codmacta & " de la variedad " & Format(Rs!codvarie, "0000")
            ElseIf Opcion = 9 Or Opcion = 10 Then
                sql = DBLet(Rs!Cuenta, "T") ' vParamAplic.CtaAboTrans
            End If
        End If
        
        
        If Opcion = 2 Or Opcion = 3 Then
            'Comprobar que ademas de existir la cuenta de ventas exista tambien
            'la cuenta ABONO ventas (sfamia.aboventa)
            '---------------------------------------------
            sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!ctaabono, "T")
'            RSconta.MoveFirst
'            RSconta.Find (SQL), , adSearchForward
'            If RSconta.EOF Then
            If Not (RegistrosAListar(sql, cConta) > 0) Then
                b = False 'no encontrado
                If Opcion = 2 Then
                    sql = Rs!ctaabono & " de la familia " & Format(Rs!codfamia, "0000")
                ElseIf Opcion = 3 Then
                    sql = Rs!ctaabono
                End If
            End If
            
            
            'comprobar cuentas alternativas solo para facturacion a CLIENTES
            '----------------------------------------------------------------
            If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
                ' Comprobar cuenta VENTA alternativa
                If DBLet(Rs!ctavent1, "T") <> "" Then
                    sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!ctavent1, "T")
'                    RSconta.MoveFirst
'                    RSconta.Find (SQL), , adSearchForward
'                    If RSconta.EOF Then
                    If Not (RegistrosAListar(sql, cConta) > 0) Then
                        b = False 'no encontrado
                        If Opcion = 2 Then
                            sql = Rs!ctavent1 & " de la familia " & Format(Rs!codfamia, "0000")
                        ElseIf Opcion = 3 Then
                            sql = Rs!ctavent1
                        End If
                    End If
                Else
                    b = False
                    sql = " o la familia no tiene asignada cuenta venta alternativa."
                End If
                
                ' Comprobar cuenta de ABONO alternativa
                If DBLet(Rs!abovent1, "T") <> "" Then
                    sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!abovent1, "T")
'                    RSconta.MoveFirst
'                    RSconta.Find (SQL), , adSearchForward
'                    If RSconta.EOF Then
                    If Not (RegistrosAListar(sql, cConta) > 0) Then
                        b = False 'no encontrado
                        If Opcion = 2 Then
                            sql = Rs!abovent1 & " de la familia " & Format(Rs!codfamia, "0000")
                        ElseIf Opcion = 3 Then
                            sql = Rs!abovent1
                        End If
                    End If
                Else
                    b = False
                    sql = " o la familia no tiene asignada cuenta abono alternativa."
                End If
            End If
            
        End If
        
        Rs.MoveNext
    Wend
    
    
'    Set RSconta = New ADODB.Recordset
'    RSconta.Open SQL, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

'    If Not RSconta.EOF Then
'        If Opcion = 1 Then
'            If cadTabla = "scafac" Then
'                'Seleccionamos los distintos clientes,cuentas que vamos a facturar
'                SQL = "SELECT DISTINCT scafac.codclien, sclien.codmacta "
'                SQL = SQL & " FROM (scafac INNER JOIN sclien ON scafac.codclien=sclien.codclien) "
'                SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
'            Else
'                'Seleccionamos los distintos proveedores,cuentas que vamos a facturar
'                SQL = "SELECT DISTINCT scafpc.codprove, sprove.codmacta "
'                SQL = SQL & " FROM (scafpc INNER JOIN sprove ON scafpc.codprove=sprove.codprove) "
'                SQL = SQL & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
'            End If

'        ElseIf Opcion = 2 Or Opcion = 3 Then
'            SQL = "SELECT distinct "
'            If Opcion = 2 Then SQL = SQL & " sartic.codfamia,"
'            If cadTabla = "scafac" Then
'                SQL = SQL & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((slifac "
'                SQL = SQL & " INNER JOIN tmpFactu ON slifac.codtipom=tmpFactu.codtipom AND slifac.numfactu=tmpFactu.numfactu AND slifac.fecfactu=tmpFactu.fecfactu) "
'                SQL = SQL & "INNER JOIN sartic ON slifac.codartic=sartic.codartic) "
'            Else
'                SQL = SQL & " sfamia.ctacompr as codmacta,sfamia.abocompr as ctaabono from ((slifpc "
'                SQL = SQL & " INNER JOIN tmpFactu ON slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu) "
'                SQL = SQL & "INNER JOIN sartic ON slifpc.codartic=sartic.codartic) "
'            End If
'            SQL = SQL & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
'        End If
        
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        b = True
'        While Not RS.EOF And b
'            SQL = "codmacta= " & DBSet(RS!Codmacta, "T")
'            RSconta.MoveFirst
'            RSconta.Find (SQL), , adSearchForward
'            If RSconta.EOF Then
'                b = False 'no encontrado
'                If Opcion = 1 Then
'                    If cadTabla = "scafac" Then
'                        SQL = RS!Codmacta & " del Cliente " & Format(RS!CodClien, "000000")
'                    Else
'                        SQL = RS!Codmacta & " del Proveedor " & Format(RS!codProve, "000000")
'                    End If
'                ElseIf Opcion = 2 Then
'                    SQL = RS!Codmacta & " de la familia " & Format(RS!codfamia, "0000")
'                ElseIf Opcion = 3 Then
'                    SQL = RS!Codmacta
'                End If
'            End If
            
'            If Opcion = 2 Then
'                'Comprobar que ademas de existir la cuenta de ventas exista tambien
'                'la cuenta ABONO ventas
'                SQL = "codmacta= " & DBSet(RS!ctaabono, "T")
'                RSconta.MoveFirst
'                RSconta.Find (SQL), , adSearchForward
'                If RSconta.EOF Then
'                    b = False 'no encontrado
'
'                    SQL = RS!ctaabono & " de la familia " & Format(RS!codfamia, "0000")
'                End If
'            End If
            
            'comprobar cuentas alternativas solo para facturacion a clientes
'            If cadTabla = "scafac" Then
'                If Opcion = 2 Then
'                    ' Comprobar cuenta venta alternativa
'                    If DBLet(RS!ctavent1, "T") <> "" Then
'                        SQL = "codmacta= " & DBSet(RS!ctavent1, "T")
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
'                            b = False 'no encontrado
'                            SQL = RS!ctavent1 & " de la familia " & Format(RS!codfamia, "0000")
'                        End If
'                    Else
'                        b = False
'                        SQL = " o la familia no tiene asignada cuenta venta alternativa."
'                    End If
'                End If
'                If Opcion = 2 Then
'                    ' Comprobar cuenta de abono alternativa
'                    If DBLet(RS!abovent1, "T") <> "" Then
'                        SQL = "codmacta= " & DBSet(RS!abovent1, "T")
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
'                            b = False 'no encontrado
'                            SQL = RS!ctaabon1 & " de la familia " & Format(RS!codfamia, "0000")
'                        End If
'                    Else
'                        b = False
'                        SQL = " o la familia no tiene asignada cuenta abono alternativa."
'                    End If
'                End If
'            End If
'            RS.MoveNext
'        Wend
'        RS.Close
'        Set RS = Nothing
        
        
        
        If Not b Then
            If Not (Opcion = 3 Or Opcion = 6 Or Opcion = 7) Then
                sql = "No existe la cta contable " & sql
            Else
                sql = "La cuenta " & sql & " no es del nivel correcto. "
                If Opcion = 3 Then sql = sql & "(Familias de art�culos)."
            End If
            sql = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & sql
            
            MsgBox sql, vbExclamation
            ComprobarCtaContable_new = False
        Else
            ComprobarCtaContable_new = True
        End If
'    Else
'        ComprobarCtaContable_new = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
    Exit Function
    
ECompCta:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function







Public Function ComprobarTiposIVA(cadTABLA As String) As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (scafac.codigiva1, codigiv2,codigiv3)
'que vamos a contabilizar existan en la contabilidad
Dim sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVA = False
    
    sql = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open sql, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        For i = 1 To 3
            If cadTABLA = "facturas" Then
                sql = "SELECT DISTINCT facturas.codiiva" & i
                sql = sql & " FROM facturas "
                sql = sql & " INNER JOIN tmpFactu ON facturas.codtipom=tmpFactu.codtipom AND facturas.numfactu=tmpFactu.numfactu AND facturas.fecfactu=tmpFactu.fecfactu "
                sql = sql & " WHERE not isnull(codiiva" & i & ")"
'                SQL = SQL & " WHERE " & " codigiv" & i & " <> 0 "
            Else
                If cadTABLA = "facturassocio" Then
                    sql = "SELECT DISTINCT facturassocio.codiiva" & i
                    sql = sql & " FROM facturassocio "
                    sql = sql & " INNER JOIN tmpFactu ON facturassocio.codtipom=tmpFactu.codtipom AND facturassocio.numfactu=tmpFactu.numfactu AND facturassocio.fecfactu=tmpFactu.fecfactu "
                    sql = sql & " WHERE not isnull(codiiva" & i & ")"
                Else
                    If cadTABLA = "scafpc" Then
                        sql = "SELECT DISTINCT scafpc.tipoiva" & i
                        sql = sql & " FROM " & cadTABLA
                        sql = sql & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
                        sql = sql & " WHERE not isnull(tipoiva" & i & ")"
        '                SQL = SQL & " WHERE " & " tipoiva" & i & " <> 0 "
                    Else
                        sql = "SELECT DISTINCT tcafpc.tipoiva" & i
                        sql = sql & " FROM " & cadTABLA
                        sql = sql & " INNER JOIN tmpFactu ON tcafpc.codtrans=tmpFactu.codtrans AND tcafpc.numfactu=tmpFactu.numfactu AND tcafpc.fecfactu=tmpFactu.fecfactu "
                        sql = sql & " WHERE not isnull(tipoiva" & i & ")"
        '                SQL = SQL & " WHERE " & " tipoiva" & i & " <> 0 "
                    
                    End If
                End If
            End If
'            SQL = SQL & " WHERE " & cadWHERE & " AND codigiv" & i & " <> 0 "

            Set Rs = New ADODB.Recordset
            Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            b = True
            While Not Rs.EOF And b
                sql = "codigiva= " & DBSet(Rs.Fields(0), "N")
                RSconta.MoveFirst
                RSconta.Find (sql), , adSearchForward
                If RSconta.EOF Then
                    b = False 'no encontrado
                    sql = "Tipo de IVA: " & Rs.Fields(0)
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        
            If Not b Then
                sql = "No existe el " & sql
                sql = "Comprobando Tipos de IVA en contabilidad..." & vbCrLf & vbCrLf & sql
            
                MsgBox sql, vbExclamation
                ComprobarTiposIVA = False
                Exit For
            Else
                ComprobarTiposIVA = True
            End If
        Next i
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompIVA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Tipo de IVA.", Err.Description
    End If
End Function


Public Function ComprobarCCoste(cadCC As String) As Boolean
Dim sql As String

    On Error GoTo ECCoste

    ComprobarCCoste = False
    sql = vUsu.Login
    If sql <> "" Then
        cadCC = DevuelveDesdeBDNew(cAgro, "straba", "codccost", "login", sql, "T")
        If cadCC <> "" Then
            'comprobar que el Centro de Coste existe en la Contabilidad
            If vParamAplic.ContabilidadNueva Then
                sql = DevuelveDesdeBDNew(cConta, "ccoste", "codccost", "codccost", cadCC, "T")
            Else
                sql = DevuelveDesdeBDNew(cConta, "cabccost", "codccost", "codccost", cadCC, "T")
            End If
            If sql <> "" Then
                ComprobarCCoste = True
            Else
                sql = "No existe el CC: " & cadCC
                sql = "Comprobando Centros de Coste en contabilidad..." & vbCrLf & vbCrLf & sql
                MsgBox sql, vbExclamation
            End If
        Else 'el usuario no tiene asignado un centro de coste
            sql = "El trabajador conectado no tiene asignado un centro de coste."
            sql = "Comprobando Centros de Coste ..." & vbCrLf & vbCrLf & sql
            MsgBox sql, vbExclamation
        End If
    End If
    
ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Centros de Coste", Err.Description
    End If
End Function


Public Function ComprobarCCoste_new(cadCC As String, cadTABLA As String, Optional Opcion As Byte) As Boolean
Dim sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECCoste

    ComprobarCCoste_new = False
    Select Case cadTABLA
        Case "facturas" ' facturas de venta
            Select Case Opcion
                Case 1
                    sql = "select distinct variedades.codccost from facturas_variedad, albaran_variedad, variedades, tmpFactu where "
                    sql = sql & " albaran_variedad.codvarie=variedades.codvarie and "
                    sql = sql & " facturas_variedad.codtipom=tmpFactu.codtipom AND facturas_variedad.numfactu=tmpFactu.numfactu AND facturas_variedad.fecfactu=tmpFactu.fecfactu and  "
                    sql = sql & " albaran_variedad.numalbar = facturas_variedad.numalbar and "
                    sql = sql & " albaran_variedad.numlinea = facturas_variedad.numlinealbar "
                Case 2
                    sql = " select distinct sfamia.codccost from facturas_envases, sartic, sfamia, tmpFactu where "
                    sql = sql & " facturas_envases.codtipom=tmpFactu.codtipom AND facturas_envases.numfactu=tmpFactu.numfactu AND facturas_envases.fecfactu=tmpFactu.fecfactu and  "
                    sql = sql & " facturas_envases.codartic = sartic.codartic and "
                    sql = sql & " sartic.codfamia = sfamia.codfamia "
                Case 3
                    If HayFacturasACuenta Then
                        sql = " select '" & vParamAplic.CCosteFraACta & "' as codccost from tmpFactu where tmpfactu.codtipom = 'EAC' "
                    Else
                        ComprobarCCoste_new = True
                        Exit Function
                    End If
            End Select
        
        Case "facturassocio" ' facturas de venta a socios
            Select Case Opcion
                Case 1
                    sql = "select distinct variedades.codccost from facturassocio_variedad, albaran_variedad, variedades, tmpFactu where "
                    sql = sql & " albaran_variedad.codvarie=variedades.codvarie and "
                    sql = sql & " facturassocio_variedad.codtipom=tmpFactu.codtipom AND facturassocio_variedad.numfactu=tmpFactu.numfactu AND facturassocio_variedad.fecfactu=tmpFactu.fecfactu and  "
                    sql = sql & " albaran_variedad.numalbar = facturassocio_variedad.numalbar and "
                    sql = sql & " albaran_variedad.numlinea = facturassocio_variedad.numlinealbar "
                Case 2
                    sql = " select distinct sfamia.codccost from facturassocio_envases, sartic, sfamia, tmpFactu where "
                    sql = sql & " facturassocio_envases.codtipom=tmpFactu.codtipom AND facturassocio_envases.numfactu=tmpFactu.numfactu AND facturassocio_envases.fecfactu=tmpFactu.fecfactu and  "
                    sql = sql & " facturassocio_envases.codartic = sartic.codartic and "
                    sql = sql & " sartic.codfamia = sfamia.codfamia "
                Case 3
                    If HayFacturasACuenta Then
                        sql = " select '" & vParamAplic.CCosteFraACta & "' as codccost from tmpFactu where tmpfactu.codtipom = 'EAC' "
                    Else
                        ComprobarCCoste_new = True
                        Exit Function
                    End If
            End Select
        
        Case "scafpc" ' facturas de compra
            sql = " select distinct sfamia.codccost from slifpc, sartic, sfamia, tmpFactu where "
            sql = sql & " slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu and  "
            sql = sql & " slifpc.codartic = sartic.codartic and "
            sql = sql & " sartic.codfamia = sfamia.codfamia "
        
        Case "tcafpc" ' facturas de transporte
            sql = "select distinct variedades.codccost from tlifpc, albaran_variedad, variedades, tmpFactu where "
            sql = sql & " albaran_variedad.codvarie=variedades.codvarie and "
            sql = sql & " tlifpc.codtrans=tmpFactu.codtrans AND tlifpc.numfactu=tmpFactu.numfactu AND tlifpc.fecfactu=tmpFactu.fecfactu and  "
            sql = sql & " albaran_variedad.numalbar = tlifpc.numalbar and "
            sql = sql & " albaran_variedad.numlinea = tlifpc.numlinea "
    
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = True

    While Not Rs.EOF And b
        If DBLet(Rs.Fields(0).Value, "T") = "" Then
            b = False
        Else
            If vParamAplic.ContabilidadNueva Then
                sql = DevuelveDesdeBDNew(cConta, "ccoste", "codccost", "codccost", Rs.Fields(0).Value, "T")
            Else
                sql = DevuelveDesdeBDNew(cConta, "cabccost", "codccost", "codccost", Rs.Fields(0).Value, "T")
            End If
            If sql = "" Then
                b = False
                Sql2 = "Centro de Coste: " & Rs.Fields(0)
            End If
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
    If Not b Then
        sql = "No existe el " & Sql2
        sql = "Comprobando Centros de Coste en contabilidad..." & vbCrLf & vbCrLf & sql
    
        MsgBox sql, vbExclamation
        ComprobarCCoste_new = False
        Exit Function
    Else
        ComprobarCCoste_new = True
    End If
    
ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Centros de Coste", Err.Description
    End If
End Function


Public Function ComprobarFormadePago(cTabla As String) As Boolean
Dim sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECCoste

    ComprobarFormadePago = False
    sql = "select distinct " & cTabla & ".codforpa from " & cTabla & ", tmpFactu where "
    sql = sql & cTabla & ".codtipom=tmpFactu.codtipom AND " & cTabla & ".numfactu=tmpFactu.numfactu AND " & cTabla & ".fecfactu=tmpFactu.fecfactu  "
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = True

    While Not Rs.EOF And b
        If Not vParamAplic.ContabilidadNueva Then
            sql = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", Rs.Fields(0).Value, "N")
        Else
            sql = DevuelveDesdeBDNew(cConta, "formapago", "codforpa", "codforpa", Rs.Fields(0).Value, "N")
        End If
        If sql = "" Then
            b = False
            Sql2 = "Formas de Pago: " & Rs.Fields(0)
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
    If Not b Then
        sql = "No existe la " & Sql2
        sql = "Comprobando Formas de Pago en contabilidad..." & vbCrLf & vbCrLf & sql
    
        MsgBox sql, vbExclamation
        ComprobarFormadePago = False
        Exit Function
    Else
        ComprobarFormadePago = True
    End If
    
ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Formas de Pago", Err.Description
    End If
End Function




Public Function PasarFactura(cadWHERE As String, CodCCost As String, CtaBan As String, cTabla As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariges.scafac --> conta.cabfact
' ariges.slifac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim sql As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    '$$$
    'Insertar en la conta Cabecera Factura
    cadMen = "Insertando Cab. Factura: " & cadMen
    b = InsertarCabFact(cTabla, cadWHERE, cadMen, vContaFra)
    
    If b Then
        CCoste = CodCCost
        'Insertar lineas de Factura en la Conta
        cadMen = "Insertando Lin. Factura: " & cadMen
        If vParamAplic.ContabilidadNueva Then
            b = InsertarLinFact_newContaNueva(cTabla, cadWHERE, cadMen)
        Else
            b = InsertarLinFact_new(cTabla, cadWHERE, cadMen)
        End If

        '++monica:a�adida la parte de insertar en tesoreria
        If b Then
            cadMen = "Insertando en Tesoreria: " & cadMen
            Select Case cTabla
                Case "facturas"
                    b = InsertarEnTesoreriaNewFac(cadWHERE, CtaBan, cadMen)
                Case "facturassocio"
                    b = InsertarEnTesoreriaNewFacSoc(cadWHERE, CtaBan, cadMen)
            End Select
        End If
        
        '++


        If b Then
            If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
        
            'Poner intconta=1 en ariagro.facturas
            cadMen = "Actualizando Factura: " & cadMen
            b = ActualizarCabFact(cTabla, cadWHERE, cadMen)
        End If
    End If
    
'    If Not b Then
'        Sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
'        Sql = Sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
'        Sql = Sql & " WHERE " & Replace(cadWhere, "facturas", "tmpFactu")
'        Conn.Execute Sql
'    End If
    
EContab:
    
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFactura = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFactura = False
        
        sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        sql = sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
        sql = sql & " WHERE " & Replace(cadWHERE, cTabla, "tmpFactu")
        conn.Execute sql
    End If
End Function


Private Function InsertarCabFact(cTabla As String, cadWHERE As String, cadErr As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Intracom As Integer
Dim SeccionHorto As Integer
Dim CadenaInsertFaclin2 As String


    On Error GoTo EInsertar
    sql = ""
    Select Case cTabla
        Case "facturas"
            sql = sql & " SELECT stipom.letraser,numfactu,fecfactu, clientes.codmacta,clientes.cliabono,year(fecfactu) as anofaccl,"
            sql = sql & "baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
            sql = sql & "totalfac,codiiva1,codiiva2,codiiva3, porcrec1, porcrec2, porcrec3, imporec1, imporec2, imporec3, clientes.codpaise "
            sql = sql & ",nomclien,domclien,codpobla,pobclien,proclien,cifclien,facturas.codforpa"
            sql = sql & " FROM (" & "facturas inner join " & "usuarios.stipom stipom on facturas.codtipom=stipom.codtipom) "
            sql = sql & "INNER JOIN " & "clientes ON facturas.codclien=clientes.codclien "
            sql = sql & " WHERE " & cadWHERE
    
        Case "facturassocio"
            SeccionHorto = DevuelveValor("select seccionhorto from rparam")
        
            sql = sql & " SELECT stipom.letraser,numfactu,fecfactu, rsocios_seccion.codmaccli codmacta,false cliabono,year(fecfactu) as anofaccl,"
            sql = sql & "baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
            sql = sql & "totalfac,codiiva1,codiiva2,codiiva3, porcrec1, porcrec2, porcrec3, imporec1, imporec2, imporec3, 0 codpaise "
            sql = sql & ",nomsocio nomclien,dirsocio domclien,codpostal codpobla,pobsocio pobclien,prosocio proclien,nifsocio cifclien,facturassocio.codforpa "
            sql = sql & " FROM ((" & "facturassocio inner join " & "usuarios.stipom stipom on facturassocio.codtipom=stipom.codtipom) "
            sql = sql & " INNER JOIN rsocios on facturassocio.codsocio = rsocios.codsocio) "
            sql = sql & "INNER JOIN " & "rsocios_seccion ON facturassocio.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(SeccionHorto, "N")
            sql = sql & " WHERE " & cadWHERE
    
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
    
        vContaFra.NumeroFactura = DBLet(Rs!NumFactu)
        vContaFra.Anofac = Year(DBLet(Rs!FecFactu))
        vContaFra.Serie = DBLet(Rs!letraser)
    
    
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = 0
        DtoGnral = 0
        BaseImp = Rs!baseimp1 + CCur(DBLet(Rs!baseimp2, "N")) + CCur(DBLet(Rs!baseimp3, "N"))
        IvaImp = DBLet(Rs!impoiva1, "N") + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
        '---- Laura 10/10/2006:  a�adir el totalfac para utilizarlo en insertar lineas
        TotalFac = Rs!TotalFac
        '----
        conCtaAlt = Rs!cliAbono

        Intracom = 0
        If Not DBSet(Rs!codpaise, "N", "S") = ValorNulo Then
            sql = ""
            sql = DevuelveDesdeBDNew(cAgro, "paises", "intracom", "codpaise", Rs!codpaise, "N")
            If sql <> "" Then Intracom = CInt(sql)
        End If
        
        If vParamAplic.ContabilidadNueva Then
            sql = ""
            sql = DBSet(Rs!letraser, "T") & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!FecFactu) & ",'FACTURACION',"
            
            ' para el caso de las rectificativas
            Dim vTipM As String
            vTipM = DevuelveValor("select codtipom from stipom where letraser = " & DBSet(Rs!letraser, "T"))
            If vTipM = "FAR" Then
                sql = sql & "'D',"
            Else
                sql = sql & "'0',"
            End If
            
            
            sql = sql & "0," & DBSet(Rs!Codforpa, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
            sql = sql & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!FecFactu, "F") & ","
            sql = sql & DBSet(Rs!Nomclien, "T") & "," & DBSet(Rs!domclien, "T") & "," & DBSet(Rs!codPobla, "T") & "," & DBSet(Rs!pobclien, "T") & ","
            sql = sql & DBSet(Rs!proclien, "T") & "," & DBSet(Rs!cifClien, "T") & ",'ES',1"
            
            cad = "(" & sql & ")"
        Else
            sql = ""
            sql = "'" & Rs!letraser & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!FecFactu) & "," & ValorNulo & ","
            sql = sql & DBSet(Rs!baseimp1, "N") & "," & DBSet(Rs!baseimp2, "N", "S") & "," & DBSet(Rs!baseimp3, "N", "S") & "," & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", "S") & "," & DBSet(Rs!porciva3, "N", "S") & ","
            sql = sql & DBSet(Rs!porcrec1, "N", "S") & "," & DBSet(Rs!porcrec2, "N", "S") & "," & DBSet(Rs!porcrec3, "N", "S") & "," & DBSet(Rs!impoiva1, "N", "N") & "," & DBSet(Rs!impoiva2, "N", "S") & "," & DBSet(Rs!impoiva3, "N", "S") & ","
            sql = sql & DBSet(Rs!imporec1, "N", "S") & "," & DBSet(Rs!imporec2, "N", "S") & "," & DBSet(Rs!imporec3, "N", "S") & ","
            sql = sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!codiiva1, "N") & "," & DBSet(Rs!codiiva2, "N", "S") & "," & DBSet(Rs!codiiva3, "N", "S") & "," & DBSet(Intracom, "N") & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & DBSet(Rs!FecFactu, "F")
            cad = "(" & sql & ")"
        End If
'        RS.MoveNext

    End If
    
    
    If vParamAplic.ContabilidadNueva Then
        sql = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
        sql = sql & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
        sql = sql & "codpais,codagente)"
        sql = sql & " VALUES " & cad
        ConnConta.Execute sql
'***
        CadenaInsertFaclin2 = ""
            
        
        'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
        'IVA 1, siempre existe
        Sql2 = "'" & Rs!letraser & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
        Sql2 = Sql2 & "1," & DBSet(Rs!baseimp1, "N") & "," & Rs!codiiva1 & "," & DBSet(Rs!porciva1, "N") & ","
        Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & ValorNulo
        CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
        
        'para las lineas
        vTipoIva(0) = Rs!codiiva1
        vPorcIva(0) = Rs!porciva1
        vPorcRec(0) = 0
        vImpIva(0) = Rs!impoiva1
        vImpRec(0) = 0
        vBaseIva(0) = Rs!baseimp1
        
        vTipoIva(1) = 0: vTipoIva(2) = 0
        
        If Not IsNull(Rs!porciva2) Then
            Sql2 = "'" & Rs!letraser & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
            Sql2 = Sql2 & "2," & DBSet(Rs!baseimp2, "N") & "," & Rs!codiiva2 & "," & DBSet(Rs!porciva2, "N") & ","
            Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva2, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
            vTipoIva(1) = Rs!codiiva2
            vPorcIva(1) = Rs!porciva2
            vPorcRec(1) = 0
            vImpIva(1) = Rs!impoiva2
            vImpRec(1) = 0
            vBaseIva(1) = Rs!baseimp2
        End If
        If Not IsNull(Rs!porciva3) Then
            Sql2 = "'" & Rs!letraser & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
            Sql2 = Sql2 & "3," & DBSet(Rs!baseimp3, "N") & "," & Rs!codiiva3 & "," & DBSet(Rs!porciva3, "N") & ","
            Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva3, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
            vTipoIva(2) = Rs!codiiva3
            vPorcIva(2) = Rs!porciva3
            vPorcRec(2) = 0
            vImpIva(2) = Rs!impoiva3
            vImpRec(2) = 0
            vBaseIva(2) = Rs!baseimp3
        End If

        sql = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
        sql = sql & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
        ConnConta.Execute sql
'***
    Else
        'Insertar en la contabilidad
        sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
        sql = sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
        sql = sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
        sql = sql & " VALUES " & cad
    
        ConnConta.Execute sql
    
    
    End If
    
    Rs.Close
    Set Rs = Nothing
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFact = False
        cadErr = Err.Description
    Else
        InsertarCabFact = True
    End If
End Function



Private Function InsertarLinFact(cadTABLA As String, cadWHERE As String, cadErr As String, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    If cadTABLA = "scafac" Then
        sql = " SELECT stipom.letraser,slifac.codtipom,numfactu,fecfactu,sartic.codfamia,sfamia.ctaventa,sfamia.ctavent1,sfamia.aboventa,sfamia.abovent1,sum(importel) as importe "
        sql = sql & " FROM ((slifac inner join usuarios.stipom stipom on slifac.codtipom=stipom.codtipom) "
        sql = sql & " inner join sartic on slifac.codartic=sartic.codartic) "
        sql = sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        sql = sql & " WHERE " & Replace(cadWHERE, "scafac", "slifac")
        sql = sql & " GROUP BY sfamia.codfamia "
    Else
        sql = " SELECT slifpc.codprove,numfactu,fecfactu,sartic.codfamia,sfamia.ctacompr,sfamia.abocompr,sum(importel) as importe "
        sql = sql & " FROM (slifpc  "
        sql = sql & " inner join sartic on slifpc.codartic=sartic.codartic) "
        sql = sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        sql = sql & " WHERE " & Replace(cadWHERE, "scafpc", "slifpc")
        sql = sql & " GROUP BY sfamia.codfamia "
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    i = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
        ImpLinea = Rs!Importe - CalcularPorcentaje(Rs!Importe, DtoPPago, 2)
        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
        ImpLinea = ImpLinea - CalcularPorcentaje(Rs!Importe, DtoGnral, 2)
        'ImpLinea = Round(ImpLinea, 2)
        '----
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        sql = ""
        Sql2 = ""
        If cadTABLA = "scafac" Then
            sql = "'" & Rs!letraser & "'," & Rs!NumFactu & "," & Year(Rs!FecFactu) & "," & i & ","
            If Not conCtaAlt Then 'cliente no tiene cuenta alternativa
                If ImpLinea >= 0 Then
                    sql = sql & DBSet(Rs!ctaventa, "T")
                Else
                    sql = sql & DBSet(Rs!aboventa, "T")
                End If
            Else
                If ImpLinea >= 0 Then
                    sql = sql & DBSet(Rs!ctavent1, "T")
                Else
                    sql = sql & DBSet(Rs!abovent1, "T")
                End If
            End If
        Else
            sql = numRegis & "," & Year(Rs!FecFactu) & "," & i & ","
            If ImpLinea >= 0 Then
                sql = sql & DBSet(Rs!ctacompr, "T")
            Else
                sql = sql & DBSet(Rs!abocompr, "T")
            End If
        End If
        Sql2 = sql & ","
        sql = sql & "," & DBSet(ImpLinea, "N") & ","
        
        If CCoste = "" Then
            sql = sql & ValorNulo
        Else
            sql = sql & DBSet(CCoste, "T")
        End If
        
        cad = cad & "(" & sql & ")" & ","
        
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)
        Sql2 = Sql2 & DBSet(totimp, "N") & ","
        If CCoste = "" Then
            Sql2 = Sql2 & ValorNulo
        Else
            Sql2 = Sql2 & DBSet(CCoste, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If


    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        If cadTABLA = "scafac" Then
            sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Else
            sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        End If
        sql = sql & " VALUES " & cad
        ConnConta.Execute sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact = False
        cadErr = Err.Description
    Else
        InsertarLinFact = True
    End If
End Function





Private Function InsertarLinFact_new(cadTABLA As String, cadWHERE As String, cadErr As String, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim CadCampo3 As String
Dim numNivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String
Dim Tipo As Byte
Dim TipoFact As String

    On Error GoTo EInLinea
    

    Select Case cadTABLA
        Case "facturas" 'VENTAS
            '[Monica]23/06/2010 Miramos que tipo de movimiento tiene la factura
            ' si es EAC (factura de anticipo) la cuenta base est� en parametros
            TipoFact = DevuelveValor("select codtipom from facturas where " & cadWHERE)
            Select Case TipoFact
                Case "EAC" ' facturas a cuenta
                    '[Monica]21/03/2013: Estas facturas tambien se contabilizan sobr la cta de la variedad si la tienen
                    'CadCampo1 = vParamAplic.CtaVentasFraACta
                    CadCampo1 = "CASE tipomer.tiptimer WHEN 0 THEN ctavtasinterior WHEN 1 THEN ctavtasexportacion WHEN 2 THEN ctavtasindustria WHEN 3 THEN ctavtasretirada WHEN 4 THEN ctavtasotros END"
                    
                    CadCampo3 = "if(facturas.codvarie is null or facturas.codtimer is null," & DBSet(vParamAplic.CtaVentasFraACta, "T") & "," & CadCampo1 & ")"
                    
                    If vEmpresa.TieneAnalitica Then
                        sql = " SELECT stipom.letraser,facturas.codtipom,numfactu,fecfactu," & CadCampo3 & " as cuenta,(baseimp1) as importe, '" & vParamAplic.CCosteFraACta & "' as codccost "
                    Else
                        sql = " SELECT stipom.letraser,facturas.codtipom,numfactu,fecfactu," & CadCampo3 & " as cuenta,(baseimp1) as importe "
                    End If
                    
                    sql = sql & " FROM (facturas inner join usuarios.stipom stipom on facturas.codtipom=stipom.codtipom "
                    sql = sql & " LEFT JOIN variedades on facturas.codvarie = variedades.codvarie) "
                    sql = sql & " LEFT JOIN tipomer on facturas.codtimer = tipomer.codtimer "
                    
                    sql = sql & " WHERE " & cadWHERE
                    If vEmpresa.TieneAnalitica Then
                        sql = sql & " GROUP BY 5,7 " '& cadCampo, codccost
                    Else
                        sql = sql & " GROUP BY 5 " '& cadCampo
                    End If
                
                Case Else
                     'comprobar si el cliente utiliza cuenta alternativa
                    If conCtaAlt Then
                        'utilizamos sfamia.ctavent1 o sfamia.abovent1
                        If TotalFac >= 0 Then
                            cadCampo = "sfamia.ctavent1"
                        Else
                            cadCampo = "sfamia.abovent1" 'si es negativa es un abono
                        End If
                    Else
                        'utilizamos sfamia.ctaventa o sfamia.aboventa
                        If TotalFac >= 0 Then
                            cadCampo = "sfamia.ctaventa"
                        Else
                            cadCampo = "sfamia.aboventa"
                        End If
                    End If
            '   select concat(raizctavtas, right(concat('000000',codvarie),5)) as cuenta from variedades
                    numNivel = DevuelveDesdeBDNew(cConta, "empresa", "numnivel", "codempre", vParamAplic.NumeroConta, "N")
                    NumDigit = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & numNivel, "codempre", vParamAplic.NumeroConta, "N")
                    NumDigit3 = DevuelveDesdeBDNew(cConta, "empresa", "numdigi3", "codempre", vParamAplic.NumeroConta, "N")
                    
            '        NumDigitAnt = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & NumNivel - 1, "codempre", vParamAplic.NumeroConta, "N")
                    
            '        CadCampo1 = "concat(concat(variedades.raizctavtas,tipomer.digicont), right(concat('0000000000',albaran_variedad.codvarie)," & (CCur(NumDigit) - CCur(NumDigit3) - 1) & "))" 'CCur(NumDigitAnt) + 1) & "))"
                    CadCampo1 = "CASE tipomer.tiptimer WHEN 0 THEN ctavtasinterior WHEN 1 THEN ctavtasexportacion WHEN 2 THEN ctavtasindustria WHEN 3 THEN ctavtasretirada WHEN 4 THEN ctavtasotros END"
                    
                    ' LINEAS DE ENVASES
                    
                    If vEmpresa.TieneAnalitica Then
                        sql = " SELECT stipom.letraser,facturas_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost "
                    Else
                        sql = " SELECT stipom.letraser,facturas_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
                    End If
                    
                    sql = sql & " FROM ((facturas_envases inner join usuarios.stipom stipom on facturas_envases.codtipom=stipom.codtipom) "
                    sql = sql & " inner join sartic on facturas_envases.codartic=sartic.codartic) "
                    sql = sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
                    sql = sql & " WHERE " & Replace(cadWHERE, "facturas", "facturas_envases")
                    If vEmpresa.TieneAnalitica Then
                        sql = sql & " GROUP BY 5,7 " '& cadCampo, codccost
                    Else
                        sql = sql & " GROUP BY 5 " '& cadCampo
                    End If
                    '[Monica]05/05/2015: la suma debe ser distinta de 0
                    sql = sql & " HAVING sum(importel) <> 0 "
                    
                    sql = sql & "Union "
                    
                    ' LINEAS DE VARIEDADES
                    
                    If vEmpresa.TieneAnalitica Then
                        sql = sql & " SELECT stipom.letraser,facturas_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe, variedades.codccost "
                    Else
                        sql = sql & " SELECT stipom.letraser,facturas_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe "
                    End If
                    sql = sql & " FROM (((((facturas_variedad inner join usuarios.stipom stipom on facturas_variedad.codtipom=stipom.codtipom) "
                    sql = sql & " inner join albaran on facturas_variedad.numalbar = albaran.numalbar) "
                    sql = sql & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
                    sql = sql & " inner join albaran_variedad on facturas_variedad.numalbar = albaran_variedad.numalbar and facturas_variedad.numlinealbar = albaran_variedad.numlinea) "
                    sql = sql & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
                    sql = sql & " WHERE " & Replace(cadWHERE, "facturas", "facturas_variedad")
                    If vEmpresa.TieneAnalitica Then
                        sql = sql & " GROUP BY 5,7 " '& cadCampo1, codccost
                    Else
                        sql = sql & " GROUP BY 5 " '& cadCampo1
                    End If
                    '[Monica]05/05/2015: la suma debe ser distinta de 0
                    sql = sql & " HAVING sum(impornet) <> 0 "
                    
                    ' LINEAS DE FACTURAS A CUENTA
                    sql = sql & "Union "
                    
'[Monica]12/03/2013: cambiamos la contabilizacion de las facturas a cuenta
'                    If vEmpresa.TieneAnalitica Then
'                        Sql = Sql & " SELECT stipom.letraser,facturas.codtipom,facturas.numfactu,facturas.fecfactu,'" & vParamAplic.CtaVentasFraACta & "' as cuenta,(baseimp1 * (-1)) as importe, '" & vParamAplic.CCosteFraACta & "' as codccost "
'                    Else
'                        Sql = Sql & " SELECT stipom.letraser,facturas.codtipom,facturas.numfactu,facturas.fecfactu,'" & vParamAplic.CtaVentasFraACta & "' as cuenta,(baseimp1 * (-1)) as importe "
'                    End If
'                    Sql = Sql & " FROM (facturas inner join usuarios.stipom stipom on facturas.codtipom=stipom.codtipom) "
'                    Sql = Sql & " INNER JOIN facturas_acuenta ON facturas.codtipom = facturas_acuenta.codtipom and facturas_acuenta.numfactu = facturas.numfactu and facturas_acuenta.fecfactu = facturas.fecfactu "
'                    Sql = Sql & " WHERE " & Replace(Replace(cadwhere, "numfactu", "facturas_acuenta.numfactu"), "fecfactu", "facturas_acuenta.fecfactu")
'                    If vEmpresa.TieneAnalitica Then
'                        Sql = Sql & " GROUP BY 5,7 " '& cadCampo1, codccost
'                    Else
'                        Sql = Sql & " GROUP BY 5 " '& cadCampo1
'                    End If
    
                    CadCampo3 = "if(facturas.codvarie is null or facturas.codtimer is null," & DBSet(vParamAplic.CtaVentasFraACta, "T") & "," & CadCampo1 & ")"
    
                    If vEmpresa.TieneAnalitica Then
                        sql = sql & " SELECT stipom.letraser,facturas_acuenta.codtipom,facturas_acuenta.numfactu,facturas_acuenta.fecfactu," & CadCampo3 & " as cuenta,(sum(baseimp1) * (-1)) as importe, variedades.codccost as codccost "
                    Else
                        sql = sql & " SELECT stipom.letraser,facturas_acuenta.codtipom,facturas_acuenta.numfactu,facturas_acuenta.fecfactu," & CadCampo3 & " as cuenta,(sum(baseimp1) * (-1)) as importe "
                    End If
                    sql = sql & " FROM (((facturas INNER JOIN facturas_acuenta ON facturas_acuenta.codtipomcta = facturas.codtipom and facturas_acuenta.numfactucta = facturas.numfactu and facturas_acuenta.fecfactucta = facturas.fecfactu) "
                    sql = sql & " LEFT JOIN variedades ON facturas.codvarie = variedades.codvarie) "
                    sql = sql & " LEFT JOIN tipomer ON facturas.codtimer = tipomer.codtimer) "
                    sql = sql & " INNER JOIN usuarios.stipom stipom ON facturas_acuenta.codtipom=stipom.codtipom"
                    sql = sql & " WHERE " & Replace(Replace(Replace(cadWHERE, "facturas", "facturas_acuenta"), "numfactu", "facturas_acuenta.numfactu"), "fecfactu", "facturas_acuenta.fecfactu")
                    
                    
                    
                    If vEmpresa.TieneAnalitica Then
                        sql = sql & " GROUP BY 5,7 " '& cadCampo1, codccost
                    Else
                        sql = sql & " GROUP BY 5 " '& cadCampo1
                    End If
                    '[Monica]05/05/2015: la suma debe ser distinta de 0
                    sql = sql & " HAVING (sum(baseimp1) * (-1)) <> 0 "
    
            End Select
        
        
        Case "facturassocio" 'VENTAS A SOCIO
             'comprobar si el cliente utiliza cuenta alternativa
            If conCtaAlt Then
                'utilizamos sfamia.ctavent1 o sfamia.abovent1
                If TotalFac >= 0 Then
                    cadCampo = "sfamia.ctavent1"
                Else
                    cadCampo = "sfamia.abovent1" 'si es negativa es un abono
                End If
            Else
                'utilizamos sfamia.ctaventa o sfamia.aboventa
                If TotalFac >= 0 Then
                    cadCampo = "sfamia.ctaventa"
                Else
                    cadCampo = "sfamia.aboventa"
                End If
            End If
    '   select concat(raizctavtas, right(concat('000000',codvarie),5)) as cuenta from variedades
            numNivel = DevuelveDesdeBDNew(cConta, "empresa", "numnivel", "codempre", vParamAplic.NumeroConta, "N")
            NumDigit = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & numNivel, "codempre", vParamAplic.NumeroConta, "N")
            NumDigit3 = DevuelveDesdeBDNew(cConta, "empresa", "numdigi3", "codempre", vParamAplic.NumeroConta, "N")
            
    '        NumDigitAnt = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & NumNivel - 1, "codempre", vParamAplic.NumeroConta, "N")
            
    '        CadCampo1 = "concat(concat(variedades.raizctavtas,tipomer.digicont), right(concat('0000000000',albaran_variedad.codvarie)," & (CCur(NumDigit) - CCur(NumDigit3) - 1) & "))" 'CCur(NumDigitAnt) + 1) & "))"
            CadCampo1 = "CASE tipomer.tiptimer WHEN 0 THEN ctavtasinterior WHEN 1 THEN ctavtasexportacion WHEN 2 THEN ctavtasindustria WHEN 3 THEN ctavtasretirada WHEN 4 THEN ctavtasotros END"
            
            ' LINEAS DE ENVASES
            
            If vEmpresa.TieneAnalitica Then
                sql = " SELECT stipom.letraser,facturassocio_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost "
            Else
                sql = " SELECT stipom.letraser,facturassocio_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
            End If
            
            sql = sql & " FROM ((facturassocio_envases inner join usuarios.stipom stipom on facturassocio_envases.codtipom=stipom.codtipom) "
            sql = sql & " inner join sartic on facturassocio_envases.codartic=sartic.codartic) "
            sql = sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
            sql = sql & " WHERE " & Replace(cadWHERE, "facturassocio", "facturassocio_envases")
            If vEmpresa.TieneAnalitica Then
                sql = sql & " GROUP BY 5,7 " '& cadCampo, codccost
            Else
                sql = sql & " GROUP BY 5 " '& cadCampo
            End If
            sql = sql & "Union "
            
            ' LINEAS DE VARIEDADES
            
            If vEmpresa.TieneAnalitica Then
                sql = sql & " SELECT stipom.letraser,facturassocio_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe, variedades.codccost "
            Else
                sql = sql & " SELECT stipom.letraser,facturassocio_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe "
            End If
            sql = sql & " FROM (((((facturassocio_variedad inner join usuarios.stipom stipom on facturassocio_variedad.codtipom=stipom.codtipom) "
            sql = sql & " inner join albaran on facturassocio_variedad.numalbar = albaran.numalbar) "
            sql = sql & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
            sql = sql & " inner join albaran_variedad on facturassocio_variedad.numalbar = albaran_variedad.numalbar and facturassocio_variedad.numlinealbar = albaran_variedad.numlinea) "
            sql = sql & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
            sql = sql & " WHERE " & Replace(cadWHERE, "facturassocio", "facturassocio_variedad")
            If vEmpresa.TieneAnalitica Then
                sql = sql & " GROUP BY 5,7 " '& cadCampo1, codccost
            Else
                sql = sql & " GROUP BY 5 " '& cadCampo1
            End If
            
        
        Case "scafpc" 'COMPRAS
            'utilizamos sfamia.ctaventa o sfamia.aboventa
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctacompr"
            Else
                cadCampo = "sfamia.abocompr"
            End If
            If vEmpresa.TieneAnalitica Then
                sql = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost"
            Else
                sql = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
            End If
            sql = sql & " FROM (slifpc  "
            sql = sql & " inner join sartic on slifpc.codartic=sartic.codartic) "
            sql = sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
            sql = sql & " WHERE " & Replace(cadWHERE, "scafpc", "slifpc")
            sql = sql & " GROUP BY " & cadCampo
            If vEmpresa.TieneAnalitica Then
                sql = sql & ", sfamia.codccost "
            End If
        Case Else ' FACTURAS DE TRANSPORTE
            'utilizamos sparam.ctaventa o sparam.aboventa
'            If TotalFac >= 0 Then
'                cadCampo = vParamAplic.CtaTraReten
'            Else
'                cadCampo = vParamAplic.CtaAboTrans
'            End If
'            Sql = " SELECT tlifpc.codtrans,numfactu,fecfactu,'" & cadCampo & "' as cuenta,sum(importel) as importe "
'            Sql = Sql & " FROM tlifpc  "
'            Sql = Sql & " WHERE " & Replace(cadWhere, "tcafpc", "tlifpc")
'            Sql = Sql & " GROUP BY '" & cadCampo & "'"

            'FACTURAS DE TRANSPORTE O DE COMISION
            sql = "select tipo from tcafpc where " & cadWHERE
            Tipo = DevuelveValor(sql) ' 0=transportista 1=comisionista
            Select Case Tipo
                Case 0 ' Transportista
        '++monica: si tipomercado = 1(exportacion) cogemos  variedades.ctatraexporta
        '          si tipomercado <> 1 (distinto de exportacion) cogemos  variedades.ctatrainterior
                    If vEmpresa.TieneAnalitica Then
                         sql = " SELECT 1, if(tipomer.tiptimer = 1,variedades.ctatraexporta,variedades.ctatrainterior) as cuenta, sum(tlifpc.importel) as importe, variedades.codccost "
                    Else
                         sql = " SELECT 1, if(tipomer.tiptimer = 1,variedades.ctatraexporta,variedades.ctatrainterior) as cuenta, sum(tlifpc.importel) as importe, '----' "
                    End If
                    sql = sql & " FROM tlifpc, albaran, albaran_variedad, variedades, tipomer"
                    sql = sql & " WHERE " & Replace(cadWHERE, "tcafpc", "tlifpc") & " and"
                    sql = sql & " tlifpc.numalbar = albaran_variedad.numalbar and "
                    sql = sql & " tlifpc.numlinea = albaran_variedad.numlinea and "
                    sql = sql & " albaran_variedad.numalbar = albaran.numalbar and "
                    sql = sql & " albaran_variedad.codvarie = variedades.codvarie and "
                    sql = sql & " albaran.codtimer = tipomer.codtimer "
                    sql = sql & " group by 1,2 "
                    sql = sql & " union "
                    sql = sql & " select 2, codmacta as cuenta, importel as importe, '----' "
                    sql = sql & " from tcafpv "
                    sql = sql & " where " & Replace(cadWHERE, "tcafpc", "tcafpv")
                    sql = sql & " group by 1,2 "
                    sql = sql & " order by 1,2 "

                Case 1 ' Comisionista
                    If vEmpresa.TieneAnalitica Then
                         sql = " SELECT 1, variedades.ctacomisionista as cuenta, sum(tlifpc.importel) as importe, variedades.codccost "
                    Else
                         sql = " SELECT 1, variedades.ctacomisionista as cuenta, sum(tlifpc.importel) as importe, '----' "
                    End If
                    sql = sql & " FROM tlifpc, albaran, albaran_variedad, variedades "
                    sql = sql & " WHERE " & Replace(cadWHERE, "tcafpc", "tlifpc") & " and"
                    sql = sql & " tlifpc.numalbar = albaran_variedad.numalbar and "
                    sql = sql & " tlifpc.numlinea = albaran_variedad.numlinea and "
                    sql = sql & " albaran_variedad.numalbar = albaran.numalbar and "
                    sql = sql & " albaran_variedad.codvarie = variedades.codvarie "
                    sql = sql & " group by 1,2 "
                    sql = sql & " order by 1,2 "
            End Select
    End Select
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    i = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
        ImpLinea = Rs!Importe - CCur(CalcularPorcentaje(Rs!Importe, DtoPPago, 2))
        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(Rs!Importe, DtoGnral, 2))
        'ImpLinea = Round(ImpLinea, 2)
        '----
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        sql = ""
        Sql2 = ""
        
        If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then 'VENTAS a clientes
            sql = "'" & Rs!letraser & "'," & Rs!NumFactu & "," & Year(Rs!FecFactu) & "," & i & ","
            sql = sql & DBSet(Rs!Cuenta, "T")
'            If Not conCtaAlt Then 'cliente no tiene cuenta alternativa
'                If ImpLinea >= 0 Then
'                    SQL = SQL & DBSet(RS!ctaventa, "T")
'                Else
'                    SQL = SQL & DBSet(RS!aboventa, "T")
'                End If
'            Else
'                If ImpLinea >= 0 Then
'                    SQL = SQL & DBSet(RS!ctavent1, "T")
'                Else
'                    SQL = SQL & DBSet(RS!abovent1, "T")
'                End If
'            End If
        Else
            If cadTABLA = "scafpc" Then 'COMPRAS
                'Laura 24/10/2006
                'SQL = numRegis & "," & Year(RS!FecFactu) & "," & i & ","
                sql = numRegis & "," & AnyoFacPr & "," & i & ","
                
    '            If ImpLinea >= 0 Then
                    sql = sql & DBSet(Rs!Cuenta, "T")
    '            Else
    '                SQL = SQL & DBSet(RS!abocompr, "T")
    '            End If
            Else 'TRANSPORTE
                sql = numRegis & "," & AnyoFacPr & "," & i & ","
                sql = sql & DBSet(Rs!Cuenta, "T")
            End If
        End If
        
        Sql2 = sql & "," 'nos guardamos la linea sin el importe por si a la �ltima hay q descontarle para q coincida con total factura
        sql = sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
            If cadTABLA = "tcafpc" Then
                If DBLet(Rs!CodCCost, "T") = "----" Then
                    sql = sql & DBSet(CCoste, "T")
                Else
                    sql = sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBLet(Rs!CodCCost, "T")
                End If
            Else
                sql = sql & DBSet(Rs!CodCCost, "T")
                CCoste = DBLet(Rs!CodCCost, "T")
            End If
        Else
            sql = sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        cad = cad & "(" & sql & ")" & ","
        
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)
        Sql2 = Sql2 & DBSet(totimp, "N") & ","
        If CCoste = "" Or CCoste = ValorNulo Then
            Sql2 = Sql2 & ValorNulo
        Else
            Sql2 = Sql2 & DBSet(CCoste, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If


    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
            sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Else
            sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        End If
        sql = sql & " VALUES " & cad
        ConnConta.Execute sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_new = False
        cadErr = Err.Description
    Else
        InsertarLinFact_new = True
    End If
End Function


Private Function InsertarLinFact_newContaNueva(cadTABLA As String, cadWHERE As String, cadErr As String, Optional numRegis As Long, Optional FraIntraCom As String, Optional FRecep As String) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim CadCampo3 As String
Dim numNivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String
Dim Tipo As Byte
Dim TipoFact As String

Dim NumeroIVA As Byte
Dim k As Integer
Dim HayQueAjustar As Boolean

Dim ImpImva As Currency
Dim ImpREC As Currency



    On Error GoTo EInLinea
    

    Select Case cadTABLA
        Case "facturas" 'VENTAS
            '[Monica]23/06/2010 Miramos que tipo de movimiento tiene la factura
            ' si es EAC (factura de anticipo) la cuenta base est� en parametros
            TipoFact = DevuelveValor("select codtipom from facturas where " & cadWHERE)
            Select Case TipoFact
                Case "EAC" ' facturas a cuenta
                    '[Monica]21/03/2013: Estas facturas tambien se contabilizan sobr la cta de la variedad si la tienen
                    'CadCampo1 = vParamAplic.CtaVentasFraACta
                    CadCampo1 = "CASE tipomer.tiptimer WHEN 0 THEN ctavtasinterior WHEN 1 THEN ctavtasexportacion WHEN 2 THEN ctavtasindustria WHEN 3 THEN ctavtasretirada WHEN 4 THEN ctavtasotros END"
                    
                    CadCampo3 = "if(facturas.codvarie is null or facturas.codtimer is null," & DBSet(vParamAplic.CtaVentasFraACta, "T") & "," & CadCampo1 & ")"
                    
                    If vEmpresa.TieneAnalitica Then
                        sql = " SELECT stipom.letraser,facturas.codtipom,numfactu,fecfactu," & CadCampo3 & " as cuenta,(baseimp1) as importe, '" & vParamAplic.CCosteFraACta & "' as codccost "
                    Else
                        sql = " SELECT stipom.letraser,facturas.codtipom,numfactu,fecfactu," & CadCampo3 & " as cuenta,(baseimp1) as importe "
                    End If
                    
                    sql = sql & ",facturas.codiiva1 codigiva, facturas.porciva1 porciva, facturas.porcrec1 porcrec "
                    
                    sql = sql & " FROM (facturas inner join usuarios.stipom stipom on facturas.codtipom=stipom.codtipom "
                    sql = sql & " LEFT JOIN variedades on facturas.codvarie = variedades.codvarie) "
                    sql = sql & " LEFT JOIN tipomer on facturas.codtimer = tipomer.codtimer "
                    
                    sql = sql & " WHERE " & cadWHERE
                    If vEmpresa.TieneAnalitica Then
                        sql = sql & " GROUP BY 5,7 " '& cadCampo, codccost
                    Else
                        sql = sql & " GROUP BY 5 " '& cadCampo
                    End If
                
                Case Else
                     'comprobar si el cliente utiliza cuenta alternativa
                    If conCtaAlt Then
                        'utilizamos sfamia.ctavent1 o sfamia.abovent1
                        If TotalFac >= 0 Then
                            cadCampo = "sfamia.ctavent1"
                        Else
                            cadCampo = "sfamia.abovent1" 'si es negativa es un abono
                        End If
                    Else
                        'utilizamos sfamia.ctaventa o sfamia.aboventa
                        If TotalFac >= 0 Then
                            cadCampo = "sfamia.ctaventa"
                        Else
                            cadCampo = "sfamia.aboventa"
                        End If
                    End If
            '   select concat(raizctavtas, right(concat('000000',codvarie),5)) as cuenta from variedades
                    numNivel = DevuelveDesdeBDNew(cConta, "empresa", "numnivel", "codempre", vParamAplic.NumeroConta, "N")
                    NumDigit = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & numNivel, "codempre", vParamAplic.NumeroConta, "N")
                    NumDigit3 = DevuelveDesdeBDNew(cConta, "empresa", "numdigi3", "codempre", vParamAplic.NumeroConta, "N")
                    
            '        NumDigitAnt = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & NumNivel - 1, "codempre", vParamAplic.NumeroConta, "N")
                    
            '        CadCampo1 = "concat(concat(variedades.raizctavtas,tipomer.digicont), right(concat('0000000000',albaran_variedad.codvarie)," & (CCur(NumDigit) - CCur(NumDigit3) - 1) & "))" 'CCur(NumDigitAnt) + 1) & "))"
                    CadCampo1 = "CASE tipomer.tiptimer WHEN 0 THEN ctavtasinterior WHEN 1 THEN ctavtasexportacion WHEN 2 THEN ctavtasindustria WHEN 3 THEN ctavtasretirada WHEN 4 THEN ctavtasotros END"
                    
                    ' LINEAS DE ENVASES
                    
                    If vEmpresa.TieneAnalitica Then
                        sql = " SELECT stipom.letraser,facturas_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost "
                    Else
                        sql = " SELECT stipom.letraser,facturas_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
                    End If
                    
                    sql = sql & ", facturas_envases.codigiva, tiposiva.porceiva porciva,  tiposiva.porcerec porcrec"
                    
                    sql = sql & " FROM (((facturas_envases inner join usuarios.stipom stipom on facturas_envases.codtipom=stipom.codtipom) "
                    sql = sql & " inner join sartic on facturas_envases.codartic=sartic.codartic) "
                    sql = sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia) "
                    sql = sql & " inner join ariconta" & vParamAplic.NumeroConta & ".tiposiva on ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = facturas_envases.codigiva"
                    sql = sql & " WHERE " & Replace(cadWHERE, "facturas", "facturas_envases")
                    If vEmpresa.TieneAnalitica Then
                        sql = sql & " GROUP BY 5,7 " '& cadCampo, codccost
                    Else
                        sql = sql & " GROUP BY 5 " '& cadCampo
                    End If
                    '[Monica]05/05/2015: la suma debe ser distinta de 0
                    sql = sql & " HAVING sum(importel) <> 0 "
                    
                    sql = sql & "Union "
                    
                    ' LINEAS DE VARIEDADES
                    
                    If vEmpresa.TieneAnalitica Then
                        sql = sql & " SELECT stipom.letraser,facturas_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe, variedades.codccost "
                    Else
                        sql = sql & " SELECT stipom.letraser,facturas_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe "
                    End If
                    sql = sql & ", facturas_variedad.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                    
                    sql = sql & " FROM (((((facturas_variedad inner join usuarios.stipom stipom on facturas_variedad.codtipom=stipom.codtipom) "
                    sql = sql & " inner join albaran on facturas_variedad.numalbar = albaran.numalbar) "
                    sql = sql & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
                    sql = sql & " inner join albaran_variedad on facturas_variedad.numalbar = albaran_variedad.numalbar and facturas_variedad.numlinealbar = albaran_variedad.numlinea) "
                    sql = sql & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
                    sql = sql & " inner join ariconta" & vParamAplic.NumeroConta & ".tiposiva on ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = facturas_variedad.codigiva "
                    sql = sql & " WHERE " & Replace(cadWHERE, "facturas", "facturas_variedad")
                    
                    If vEmpresa.TieneAnalitica Then
                        sql = sql & " GROUP BY 5,7,8 " '& cadCampo1, codccost, codigiva
                    Else
                        sql = sql & " GROUP BY 5,7  " '& cadCampo1, codigiva
                    End If
                    '[Monica]05/05/2015: la suma debe ser distinta de 0
                    sql = sql & " HAVING sum(impornet) <> 0 "
                    
                    ' LINEAS DE FACTURAS A CUENTA
                    sql = sql & "Union "
                    
'[Monica]12/03/2013: cambiamos la contabilizacion de las facturas a cuenta
'                    If vEmpresa.TieneAnalitica Then
'                        Sql = Sql & " SELECT stipom.letraser,facturas.codtipom,facturas.numfactu,facturas.fecfactu,'" & vParamAplic.CtaVentasFraACta & "' as cuenta,(baseimp1 * (-1)) as importe, '" & vParamAplic.CCosteFraACta & "' as codccost "
'                    Else
'                        Sql = Sql & " SELECT stipom.letraser,facturas.codtipom,facturas.numfactu,facturas.fecfactu,'" & vParamAplic.CtaVentasFraACta & "' as cuenta,(baseimp1 * (-1)) as importe "
'                    End If
'                    Sql = Sql & " FROM (facturas inner join usuarios.stipom stipom on facturas.codtipom=stipom.codtipom) "
'                    Sql = Sql & " INNER JOIN facturas_acuenta ON facturas.codtipom = facturas_acuenta.codtipom and facturas_acuenta.numfactu = facturas.numfactu and facturas_acuenta.fecfactu = facturas.fecfactu "
'                    Sql = Sql & " WHERE " & Replace(Replace(cadwhere, "numfactu", "facturas_acuenta.numfactu"), "fecfactu", "facturas_acuenta.fecfactu")
'                    If vEmpresa.TieneAnalitica Then
'                        Sql = Sql & " GROUP BY 5,7 " '& cadCampo1, codccost
'                    Else
'                        Sql = Sql & " GROUP BY 5 " '& cadCampo1
'                    End If
    
                    CadCampo3 = "if(facturas.codvarie is null or facturas.codtimer is null," & DBSet(vParamAplic.CtaVentasFraACta, "T") & "," & CadCampo1 & ")"
    
                    If vEmpresa.TieneAnalitica Then
                        sql = sql & " SELECT stipom.letraser,facturas_acuenta.codtipom,facturas_acuenta.numfactu,facturas_acuenta.fecfactu," & CadCampo3 & " as cuenta,(sum(baseimp1) * (-1)) as importe, variedades.codccost as codccost "
                    Else
                        sql = sql & " SELECT stipom.letraser,facturas_acuenta.codtipom,facturas_acuenta.numfactu,facturas_acuenta.fecfactu," & CadCampo3 & " as cuenta,(sum(baseimp1) * (-1)) as importe "
                    End If
                    sql = sql & ", facturas.codiiva1 codigiva, facturas.porciva1 porciva, facturas.porcrec1 porcrec "
                    
                    sql = sql & " FROM (((facturas INNER JOIN facturas_acuenta ON facturas_acuenta.codtipomcta = facturas.codtipom and facturas_acuenta.numfactucta = facturas.numfactu and facturas_acuenta.fecfactucta = facturas.fecfactu) "
                    sql = sql & " LEFT JOIN variedades ON facturas.codvarie = variedades.codvarie) "
                    sql = sql & " LEFT JOIN tipomer ON facturas.codtimer = tipomer.codtimer) "
                    sql = sql & " INNER JOIN usuarios.stipom stipom ON facturas_acuenta.codtipom=stipom.codtipom"
                    sql = sql & " WHERE " & Replace(Replace(Replace(cadWHERE, "facturas", "facturas_acuenta"), "numfactu", "facturas_acuenta.numfactu"), "fecfactu", "facturas_acuenta.fecfactu")
                    
                    If vEmpresa.TieneAnalitica Then
                        sql = sql & " GROUP BY 5,7,8 " '& cadCampo1, codccost, codigiva
                    Else
                        sql = sql & " GROUP BY 5,7 " '& cadCampo1, codigiva
                    End If
                    
                    '[Monica]05/05/2015: la suma debe ser distinta de 0
                    sql = sql & " HAVING (sum(baseimp1) * (-1)) <> 0 "
    
            End Select
        
        
        Case "facturassocio" 'VENTAS A SOCIO
             'comprobar si el cliente utiliza cuenta alternativa
            If conCtaAlt Then
                'utilizamos sfamia.ctavent1 o sfamia.abovent1
                If TotalFac >= 0 Then
                    cadCampo = "sfamia.ctavent1"
                Else
                    cadCampo = "sfamia.abovent1" 'si es negativa es un abono
                End If
            Else
                'utilizamos sfamia.ctaventa o sfamia.aboventa
                If TotalFac >= 0 Then
                    cadCampo = "sfamia.ctaventa"
                Else
                    cadCampo = "sfamia.aboventa"
                End If
            End If
    '   select concat(raizctavtas, right(concat('000000',codvarie),5)) as cuenta from variedades
            numNivel = DevuelveDesdeBDNew(cConta, "empresa", "numnivel", "codempre", vParamAplic.NumeroConta, "N")
            NumDigit = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & numNivel, "codempre", vParamAplic.NumeroConta, "N")
            NumDigit3 = DevuelveDesdeBDNew(cConta, "empresa", "numdigi3", "codempre", vParamAplic.NumeroConta, "N")
            
    '        NumDigitAnt = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & NumNivel - 1, "codempre", vParamAplic.NumeroConta, "N")
            
    '        CadCampo1 = "concat(concat(variedades.raizctavtas,tipomer.digicont), right(concat('0000000000',albaran_variedad.codvarie)," & (CCur(NumDigit) - CCur(NumDigit3) - 1) & "))" 'CCur(NumDigitAnt) + 1) & "))"
            CadCampo1 = "CASE tipomer.tiptimer WHEN 0 THEN ctavtasinterior WHEN 1 THEN ctavtasexportacion WHEN 2 THEN ctavtasindustria WHEN 3 THEN ctavtasretirada WHEN 4 THEN ctavtasotros END"
            
            ' LINEAS DE ENVASES
            
            If vEmpresa.TieneAnalitica Then
                sql = " SELECT stipom.letraser,facturassocio_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost "
            Else
                sql = " SELECT stipom.letraser,facturassocio_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
            End If
            
            sql = sql & ", facturassocio_envases.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
            
            sql = sql & " FROM (((facturassocio_envases inner join usuarios.stipom stipom on facturassocio_envases.codtipom=stipom.codtipom) "
            sql = sql & " inner join sartic on facturassocio_envases.codartic=sartic.codartic) "
            sql = sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia) "
            sql = sql & " inner join ariconta" & vParamAplic.NumeroConta & ".tiposiva on ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = facturassocio_envases.codigiva "
            sql = sql & " WHERE " & Replace(cadWHERE, "facturassocio", "facturassocio_envases")
            If vEmpresa.TieneAnalitica Then
                sql = sql & " GROUP BY 5,7 " '& cadCampo, codccost
            Else
                sql = sql & " GROUP BY 5 " '& cadCampo
            End If
            sql = sql & "Union "
            
            ' LINEAS DE VARIEDADES
            
            If vEmpresa.TieneAnalitica Then
                sql = sql & " SELECT stipom.letraser,facturassocio_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe, variedades.codccost "
            Else
                sql = sql & " SELECT stipom.letraser,facturassocio_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe "
            End If
            sql = sql & ", facturassocio_variedad.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
            
            sql = sql & " FROM (((((facturassocio_variedad inner join usuarios.stipom stipom on facturassocio_variedad.codtipom=stipom.codtipom) "
            sql = sql & " inner join albaran on facturassocio_variedad.numalbar = albaran.numalbar) "
            sql = sql & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
            sql = sql & " inner join albaran_variedad on facturassocio_variedad.numalbar = albaran_variedad.numalbar and facturassocio_variedad.numlinealbar = albaran_variedad.numlinea) "
            sql = sql & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
            sql = sql & " inner join ariconta" & vParamAplic.NumeroConta & ".tiposiva on ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = facturassocio_variedad.codigiva "
            
            sql = sql & " WHERE " & Replace(cadWHERE, "facturassocio", "facturassocio_variedad")
            If vEmpresa.TieneAnalitica Then
                sql = sql & " GROUP BY 5,7, 8 " '& cadCampo1, codccost, codigiva
            Else
                sql = sql & " GROUP BY 5,7 " '& cadCampo1, codigiva
            End If
            
        
        Case "scafpc" 'COMPRAS
            'utilizamos sfamia.ctaventa o sfamia.aboventa
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctacompr"
            Else
                cadCampo = "sfamia.abocompr"
            End If
            If vEmpresa.TieneAnalitica Then
                sql = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost"
                If FraIntraCom <> "" Then
                    sql = sql & "," & DBSet(FraIntraCom, "N") & ",sartic.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                Else
                    sql = sql & ",sartic.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                End If
            Else
                sql = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
                If FraIntraCom <> "" Then
                    sql = sql & "," & DBSet(FraIntraCom, "N") & ",sartic.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                Else
                    sql = sql & ",sartic.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                End If
            End If
            sql = sql & " FROM ((slifpc  "
            sql = sql & " inner join sartic on slifpc.codartic=sartic.codartic) "
            sql = sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia) "
            sql = sql & " inner join ariconta" & vParamAplic.NumeroConta & ".tiposiva on ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = sartic.codigiva "
            sql = sql & " WHERE " & Replace(cadWHERE, "scafpc", "slifpc")
            sql = sql & " GROUP BY " & cadCampo
            If vEmpresa.TieneAnalitica Then
                sql = sql & ", sfamia.codccost "
            End If
        Case Else ' FACTURAS DE TRANSPORTE
            'utilizamos sparam.ctaventa o sparam.aboventa
'            If TotalFac >= 0 Then
'                cadCampo = vParamAplic.CtaTraReten
'            Else
'                cadCampo = vParamAplic.CtaAboTrans
'            End If
'            Sql = " SELECT tlifpc.codtrans,numfactu,fecfactu,'" & cadCampo & "' as cuenta,sum(importel) as importe "
'            Sql = Sql & " FROM tlifpc  "
'            Sql = Sql & " WHERE " & Replace(cadWhere, "tcafpc", "tlifpc")
'            Sql = Sql & " GROUP BY '" & cadCampo & "'"

            'FACTURAS DE TRANSPORTE O DE COMISION
            sql = "select tipo from tcafpc where " & cadWHERE
            Tipo = DevuelveValor(sql) ' 0=transportista 1=comisionista
            Select Case Tipo
                Case 0 ' Transportista
        '++monica: si tipomercado = 1(exportacion) cogemos  variedades.ctatraexporta
        '          si tipomercado <> 1 (distinto de exportacion) cogemos  variedades.ctatrainterior
       
                    If vEmpresa.TieneAnalitica Then
                         sql = " SELECT 1, if(tipomer.tiptimer = 1,variedades.ctatraexporta,variedades.ctatrainterior) as cuenta, sum(tlifpc.importel) as importe, variedades.codccost "
                         sql = sql & ", " & vParamAplic.CodIvaTrans & " codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                    Else
                         sql = " SELECT 1, if(tipomer.tiptimer = 1,variedades.ctatraexporta,variedades.ctatrainterior) as cuenta, sum(tlifpc.importel) as importe, '----' "
                         sql = sql & ", " & vParamAplic.CodIvaTrans & " codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec"
                    End If
                    sql = sql & " FROM tlifpc, albaran, albaran_variedad, variedades, tipomer, ariconta" & vParamAplic.NumeroConta & ".tiposiva "
                    sql = sql & " WHERE " & Replace(cadWHERE, "tcafpc", "tlifpc") & " and"
                    sql = sql & " tlifpc.numalbar = albaran_variedad.numalbar and "
                    sql = sql & " tlifpc.numlinea = albaran_variedad.numlinea and "
                    sql = sql & " albaran_variedad.numalbar = albaran.numalbar and "
                    sql = sql & " albaran_variedad.codvarie = variedades.codvarie and "
                    sql = sql & " albaran.codtimer = tipomer.codtimer "
                    sql = sql & " and ariconta" & vParamAplic.NumeroConta & ".tiposiva = " & DBSet(vParamAplic.CodIvaTrans, "N")
                    sql = sql & " group by 1,2,4,5,6,7 "
                    sql = sql & " union "
                    sql = sql & " select 2, codmacta as cuenta, importel as importe, '----' "
                    sql = sql & ", " & vParamAplic.CodIvaTrans & " codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                    sql = sql & " from tcafpv "
                    sql = sql & " where " & Replace(cadWHERE, "tcafpc", "tcafpv")
                    sql = sql & " and ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = " & DBSet(vParamAplic.CodIvaTrans, "N")
                    sql = sql & " group by 1,2,4,5,6,7 "
                    sql = sql & " order by 1,2 "

                Case 1 ' Comisionista
                    If vEmpresa.TieneAnalitica Then
                         sql = " SELECT 1, variedades.ctacomisionista as cuenta, sum(tlifpc.importel) as importe, variedades.codccost "
                         sql = sql & ", " & vParamAplic.CodIvaTrans & " codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                    Else
                         sql = " SELECT 1, variedades.ctacomisionista as cuenta, sum(tlifpc.importel) as importe, '----' "
                         sql = sql & ", " & vParamAplic.CodIvaTrans & " codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                    End If
                    sql = sql & " FROM tlifpc, albaran, albaran_variedad, variedades  "
                    sql = sql & " WHERE " & Replace(cadWHERE, "tcafpc", "tlifpc") & " and"
                    sql = sql & " tlifpc.numalbar = albaran_variedad.numalbar and "
                    sql = sql & " tlifpc.numlinea = albaran_variedad.numlinea and "
                    sql = sql & " albaran_variedad.numalbar = albaran.numalbar and "
                    sql = sql & " albaran_variedad.codvarie = variedades.codvarie "
                    sql = sql & " and ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = " & DBSet(vParamAplic.CodIvaTrans, "N")
                    sql = sql & " group by 1,2,4,5,6,7 "
                    sql = sql & " order by 1,2 "
            End Select
    End Select
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenKeyset, adLockPessimistic, adCmdText

    cad = ""
    i = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
        ImpLinea = Rs!Importe - CCur(CalcularPorcentaje(Rs!Importe, DtoPPago, 2))
        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(Rs!Importe, DtoGnral, 2))
        'ImpLinea = Round(ImpLinea, 2)
        '----
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        sql = ""
        Sql2 = ""
        
        If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then 'VENTAS a clientes
            sql = "'" & Rs!letraser & "'," & Rs!NumFactu & "," & Year(Rs!FecFactu) & "," & i & ","
            sql = sql & DBSet(Rs!Cuenta, "T") & ","
        Else
            If cadTABLA = "scafpc" Then 'COMPRAS
                'Laura 24/10/2006
                'SQL = numRegis & "," & Year(RS!FecFactu) & "," & i & ","
                sql = DBSet(SerieFraPro, "T") & "," & numRegis & "," & DBSet(FRecep, "F") & "," & AnyoFacPr & "," & i & ","
                sql = sql & DBSet(Rs!Cuenta, "T") & ","
            
            Else 'TRANSPORTE
                sql = DBSet(SerieFraPro, "T") & "," & numRegis & "," & DBSet(FRecep, "F") & "," & AnyoFacPr & "," & i & ","
                sql = sql & DBSet(Rs!Cuenta, "T") & ","
            End If
        End If
        
        'Vemos que tipo de IVA es en el vector de importes
        NumeroIVA = 127
        For k = 0 To 2
            If Rs!Codigiva = vTipoIva(k) Then
                NumeroIVA = k
                Exit For
            End If
        Next
        If NumeroIVA > 100 Then Err.Raise 513, "Error obteniendo IVA: " & Rs!Codigiva
        
        
        If vEmpresa.TieneAnalitica Then
            If cadTABLA = "tcafpc" Then
                If DBLet(Rs!CodCCost, "T") = "----" Then
                    sql = sql & DBSet(CCoste, "T")
                Else
                    sql = sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBLet(Rs!CodCCost, "T")
                End If
            Else
                sql = sql & DBSet(Rs!CodCCost, "T")
                CCoste = DBLet(Rs!CodCCost, "T")
            End If
        Else
            sql = sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
            sql = sql & "," & DBSet(Rs!FecFactu, "F")
        End If
        
        vBaseIva(NumeroIVA) = vBaseIva(NumeroIVA) - ImpLinea   'Para ajustar el importe y que no haya descuadre
        HayQueAjustar = False
        If vBaseIva(NumeroIVA) <> 0 Then
            'falta importe.
            'Puede ser que hayan mas lineas, o haya descuadre. Como esta ordenado por tipo de iva
            Rs.MoveNext
            If Rs.EOF Then
                'No hay mas lineas
                'Hay que ajustar SI o SI
                HayQueAjustar = True
            Else
                'Si que hay mas lineas.
                'Son del mismo tipo de IVA
                If Rs!Codigiva <> vTipoIva(0) Then
                    'NO es el mismo tipo de IVA
                    'Hay que ajustar
                    HayQueAjustar = True
                End If
            End If
            Rs.MovePrevious
        End If
        
        sql = sql & "," & vTipoIva(NumeroIVA) & "," & DBSet(vPorcIva(NumeroIVA), "N") & "," & DBSet(vPorcRec(NumeroIVA), "N", "S") & ","
        
        If HayQueAjustar Then
            Stop
        Else
        
        End If

        
        'Caluclo el importe de IVA y el de recargo de equivalencia
        ImpImva = vPorcIva(NumeroIVA) / 100
        ImpImva = Round2(ImpLinea * ImpImva, 2)
        If vPorcRec(NumeroIVA) = 0 Then
            ImpREC = 0
        Else
            ImpREC = vPorcRec(NumeroIVA) / 100
            ImpREC = Round2(ImpLinea * ImpREC, 2)
        End If
        vImpIva(NumeroIVA) = vImpIva(NumeroIVA) - ImpImva
        vImpRec(NumeroIVA) = vImpRec(NumeroIVA) - ImpREC
        
        
        ' baseimpo , impoiva, imporec, aplicret, CodCCost
        sql = sql & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpREC, "N", "S")
        
        ' si la linea lleva retencion
        If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then 'VENTAS a clientes
        Else
            sql = sql & ",0"
        End If
      
'        Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la �ltima hay q descontarle para q coincida con total factura
'        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        cad = cad & "(" & sql & ")" & ","
        
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
'    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
'    'de la factura
'    If totimp <> BaseImp Then
''        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
'        'en SQL esta la ult linea introducida
'        totimp = BaseImp - totimp
'        totimp = ImpLinea + totimp '(+- diferencia)
'        Sql2 = Sql2 & DBSet(totimp, "N") & ","
'        If CCoste = "" Or CCoste = ValorNulo Then
'            Sql2 = Sql2 & ValorNulo
'        Else
'            Sql2 = Sql2 & DBSet(CCoste, "T")
'        End If
'        If SQLaux <> "" Then 'hay mas de una linea
'            cad = SQLaux & "(" & Sql2 & ")" & ","
'        Else 'solo una linea
'            cad = "(" & Sql2 & ")" & ","
'        End If
'
''        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
''        cad = Replace(cad, SQL, Aux)
'    End If


    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
            sql = "INSERT INTO factcli_lineas (numserie,numfactu,anofactu,numlinea,codmacta,codccost,fecfactu,codigiva,porciva,porcrec,baseimpo,impoiva,imporec) "
        Else
            sql = "INSERT INTO factpro_lineas (numserie,numregis,fecharec,anofactu,numlinea,codmacta,codccost,codigiva,porciva,porcrec,baseimpo,impoiva,imporec,aplicret) "
        End If
        sql = sql & " VALUES " & cad
        ConnConta.Execute sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_newContaNueva = False
        cadErr = Err.Description
    Else
        InsertarLinFact_newContaNueva = True
    End If
End Function



Private Function ActualizarCabFact(cadTABLA As String, cadWHERE As String, cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim sql As String

    On Error GoTo EActualizar
    
    sql = "UPDATE " & cadTABLA & " SET intconta=1 "
    sql = sql & " WHERE " & cadWHERE

    conn.Execute sql
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCabFact = False
        cadErr = Err.Description
    Else
        ActualizarCabFact = True
    End If
End Function



'----------------------------------------------------------------------
' FACTURAS PROVEEDOR
'----------------------------------------------------------------------

Public Function PasarFacturaProv(cadWHERE As String, CodCCost As String, FechaFin As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariges.scafpc --> conta.cabfactprov
' ariges.slifpc --> conta.linfactprov
'Actualizar la tabla ariges.scafpc.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim sql As String
Dim Mc As Contadores
Dim FraIntraCom2 As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
        
    
    Set Mc = New Contadores
    
    '---- Insertar en la conta Cabecera Factura
    b = InsertarCabFactProv(cadWHERE, cadMen, Mc, FechaFin, vContaFra, FraIntraCom2)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        CCoste = CodCCost
        '---- Insertar lineas de Factura en la Conta
        If Not vParamAplic.ContabilidadNueva Then
            b = InsertarLinFact_new("scafpc", cadWHERE, cadMen, Mc.Contador)
        Else
            b = InsertarLinFact_newContaNueva("scafpc", cadWHERE, cadMen, Mc.Contador, FraIntraCom2)
        End If
        cadMen = "Insertando Lin. Factura: " & cadMen

        If b Then
            If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaProv(vContaFra.NumeroFactura, vContaFra.Anofac)
            
            '---- Poner intconta=1 en ariges.scafac
            b = ActualizarCabFact("scafpc", cadWHERE, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
'    If Not b Then
'        SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
'        SQL = SQL & " Select *," & DBSet(Mid(cadMen, 1, 200), "T") & " as error From tmpFactu "
'        SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
'        Conn.Execute SQL
'    End If
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaProv = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaProv = False
        If Not b Then
            InsertarTMPErrFac cadMen, cadWHERE
'            SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
'            SQL = SQL & " Select *," & DBSet(Mid(cadMen, 1, 200), "T") & " as error From tmpFactu "
'            SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
'            Conn.Execute SQL
        End If
    End If
End Function


Private Function InsertarCabFactProv(cadWHERE As String, cadErr As String, ByRef Mc As Contadores, FechaFin As String, ByRef vCF As cContabilizarFacturas, ByRef EsFacturaIntracom2 As String) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el a�o de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Intracom As Integer

Dim TipoOpera As Byte
Dim CadenaInsertFaclin2     As String
Dim ImporAux As Currency

Dim Aux As String
Dim Sql2 As String

    On Error GoTo EInsertar
       
    
    sql = " SELECT fecfactu,year(fecrecep) as anofacpr,fecrecep,numfactu,proveedor.codmacta,"
    sql = sql & "scafpc.dtoppago,scafpc.dtognral,baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    sql = sql & "totalfac,tipoiva1,tipoiva2,tipoiva3,proveedor.codprove, proveedor.nomprove, proveedor.tipprove, scafpc.confacpr "
    sql = sql & ",scafpc.domprove,scafpc.codpobla,scafpc.pobprove,scafpc.proprove,scafpc.nifprove,scafpc.codforpa "
    sql = sql & " FROM " & "scafpc "
    sql = sql & "INNER JOIN " & "proveedor ON scafpc.codprove=proveedor.codprove "
    sql = sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
    
        If Mc.ConseguirContador("1", (Rs!FecRecep <= CDate(FechaFin) - 365), True) = 0 Then
        
            vCF.NumeroFactura = Mc.Contador
            vCF.Anofac = Year(DBLet(Rs!FecRecep))
            vvFecRecep = DBLet(Rs!FecRecep)
            
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            DtoPPago = Rs!DtoPPago
            DtoGnral = Rs!DtoGnral
            BaseImp = Rs!BaseIVA1 + CCur(DBLet(Rs!BaseIVA2, "N")) + CCur(DBLet(Rs!BaseIVA3, "N"))
            TotalFac = Rs!TotalFac
            AnyoFacPr = Rs!anofacpr
            
            Intracom = DBLet(Rs!tipprove, "N")
            If Intracom = 2 Then Intracom = 0
            
            Nulo2 = "N"
            Nulo3 = "N"
            If DBLet(Rs!BaseIVA2, "N") = "0" Then Nulo2 = "S"
            If DBLet(Rs!BaseIVA3, "N") = "0" Then Nulo3 = "S"
            sql = ""
            If vParamAplic.ContabilidadNueva Then sql = "'" & SerieFraPro & "',"
            
            sql = sql & Mc.Contador & "," & DBSet(Rs!FecFactu, "F") & "," & Rs!anofacpr & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!NumFactu, "T") & "," & DBSet(Rs!Codmacta, "T") & "," & DBSet(Rs!confacpr, "T") & ","
            
            If Not vParamAplic.ContabilidadNueva Then
                sql = sql & DBSet(Rs!BaseIVA1, "N") & "," & DBSet(Rs!BaseIVA2, "N", "S") & "," & DBSet(Rs!BaseIVA3, "N", "S") & ","
                sql = sql & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", Nulo2) & "," & DBSet(Rs!porciva3, "N", Nulo3) & ","
                sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & DBSet(Rs!impoiva2, "N", Nulo2) & "," & DBSet(Rs!impoiva3, "N", Nulo3) & ","
                sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                sql = sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!TipoIVA1, "N") & "," & DBSet(Rs!TipoIVA2, "N", Nulo2) & "," & DBSet(Rs!TipoIVA3, "N", Nulo3) & "," & DBSet(Intracom, "N") & ","
            Else
                sql = sql & DBSet(Rs!nomprove, "T") & "," & DBSet(Rs!domprove, "T", "S") & ","
                sql = sql & DBSet(Rs!codPobla, "T", "S") & "," & DBSet(Rs!pobprove, "T", "S") & "," & DBSet(Rs!proprove, "T", "S") & ","
                sql = sql & DBSet(Rs!nifProve, "F", "S") & ",'ES',"
                sql = sql & DBSet(Rs!Codforpa, "N") & ","
                
                TipoOpera = 0
                 'IVA ES CERO
                If Rs!tipprove = 1 Then
                    'intracomunitaria
                    TipoOpera = 1
                Else
                    'Exstranjero
                     If Rs!tipprove = 1 Then TipoOpera = 2
                End If
                
                Aux = "0"
                Select Case TipoOpera
                Case 0
                    If Rs!TotalFac < 0 Then
                        Aux = "D"
                    Else
                        If Not IsNull(Rs!TipoIVA2) Then Aux = "C"
                    End If
                
                Case 1
                    Aux = "P"
                
                Case 4
                    Aux = "I"
                End Select
                
                'codopera,codconce340,codintra
                sql = sql & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                
                
                'para las lineas
                'factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
                'IVA 1, siempre existe
                Aux = "'" & SerieFraPro & "'," & Mc.Contador & "," & DBSet(Rs!FecRecep, "F") & "," & Rs!anofacpr & ","
                
                Sql2 = Aux & "1," & DBSet(Rs!BaseIVA1, "N") & "," & Rs!TipoIVA1 & "," & DBSet(Rs!porciva1, "N") & ","
                Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
                vTipoIva(0) = Rs!TipoIVA1
                vPorcIva(0) = Rs!porciva1
                vPorcRec(0) = 0
                vImpIva(0) = Rs!impoiva1
                vImpRec(0) = 0
                vBaseIva(0) = Rs!BaseIVA1
                
                vTipoIva(1) = 0: vTipoIva(2) = 0
                
                If Not IsNull(Rs!porciva2) Then
                    Sql2 = Aux & "2," & DBSet(Rs!BaseIVA2, "N") & "," & Rs!TipoIVA2 & "," & DBSet(Rs!porciva2, "N") & ","
                    Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva2, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                    vTipoIva(1) = Rs!TipoIVA2
                    vPorcIva(1) = Rs!porciva2
                    vPorcRec(1) = 0
                    vImpIva(1) = Rs!impoiva2
                    vImpRec(1) = 0
                    vBaseIva(1) = Rs!BaseIVA2
                
                End If
                If Not IsNull(Rs!porciva3) Then
                    Sql2 = Aux & "3," & DBSet(Rs!BaseIVA3, "N") & "," & Rs!TipoIVA3 & "," & DBSet(Rs!porciva3, "N") & ","
                    Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva3, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                    vTipoIva(2) = Rs!TipoIVA3
                    vPorcIva(2) = Rs!porciva3
                    vPorcRec(2) = 0
                    vImpIva(2) = Rs!impoiva3
                    vImpRec(2) = 0
                    vBaseIva(2) = Rs!BaseIVA3
                End If
                
                    
                    
                'Los totales
                'totbases,totbasesret,totivas,totrecargo,totfacpr,
                ImporAux = Rs!BaseIVA1 + DBLet(Rs!BaseIVA2, "N") + DBLet(Rs!BaseIVA3, "N")
                sql = sql & DBSet(ImporAux, "N") & "," & ValorNulo & ","
                'totivas
                ImporAux = Rs!impoiva1 + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
                sql = sql & DBSet(ImporAux, "N") & "," & DBSet(Rs!TotalFac, "N") & ","
                        
                  
                EsFacturaIntracom2 = ""
                If DBLet(Rs!tipprove, "N") = 1 Then
                    'OK es intracomunitaria
                    EsFacturaIntracom2 = Rs!TipoIVA1
                End If
            
            End If
           
            'datos de retencion
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            If vParamAplic.ContabilidadNueva Then sql = sql & "0"
            
            ' Antigua: numdiari,fechaent,numasien,nodeducible)
            If Not vParamAplic.ContabilidadNueva Then sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
            
            cad = cad & "(" & sql & ")"
            
            If vParamAplic.ContabilidadNueva Then
                sql = "INSERT INTO factpro(numserie,numregis,fecfactu,anofactu,fecharec,fecliqpr,numfactu,codmacta,observa,nommacta,"
                sql = sql & "dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,codforpa,codopera,codconce340,codintra,"
                sql = sql & "totbases,totbasesret,totivas,totfacpr,retfacpr , trefacpr, cuereten, tiporeten)"
                sql = sql & " VALUES " & cad
                ConnConta.Execute sql
            Else
                'Insertar en la contabilidad
                sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,fecliqpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                sql = sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                sql = sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,nodeducible) "
                sql = sql & " VALUES " & cad
                ConnConta.Execute sql
            End If
            
            If vParamAplic.ContabilidadNueva Then
                'Las  lineas de IVA
                sql = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
                sql = sql & " VALUES " & CadenaInsertFaclin2
                ConnConta.Execute sql
            End If
      
            
            'a�adido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            sql = sql & ",'" & DevNombreSQL(Rs!NumFactu) & " @ " & Format(Rs!FecFactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!nomprove) & "'," & Rs!codProve & ")"
            conn.Execute sql
            
            
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactProv = False
        cadErr = Err.Description
    Else
        InsertarCabFactProv = True
    End If
End Function



Public Sub FechasEjercicioConta(FIni As String, FFin As String)
'Dim RS As ADODB.Recordset
'
'    On Error GoTo EFechas
'
'    FIni = "Select fechaini,fechafin From parametros"
'    Set RS = New ADODB.Recordset
'    RS.Open FIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not RS.EOF Then
'        FIni = DBLet(RS!FechaIni, "F")
'        FFin = DBLet(RS!FechaFin, "F")
'    End If
'    RS.Close
'    Set RS = Nothing
'
'EFechas:
'    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------------
' FACTURAS TRANSPORTE
'----------------------------------------------------------------------

Public Function PasarFacturaTrans(cadWHERE As String, CodCCost As String, FechaFin As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariagro.tcafpc --> conta.cabfactprov
' ariagro.tlifpc --> conta.linfactprov
'Actualizar la tabla ariges.scafpc.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim sql As String
Dim Mc As Contadores


    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
        
    
    Set Mc = New Contadores
    
    '---- Insertar en la conta Cabecera Factura
    b = InsertarCabFactTrans(cadWHERE, cadMen, Mc, FechaFin, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        CCoste = CodCCost
        '---- Insertar lineas de Factura en la Conta
        If Not vParamAplic.ContabilidadNueva Then
            b = InsertarLinFact_new("tcafpc", cadWHERE, cadMen, Mc.Contador)
        Else
            b = InsertarLinFact_newContaNueva("tcafpc", cadWHERE, cadMen, Mc.Contador)
        
        End If
        cadMen = "Insertando Lin. Factura: " & cadMen

        If b Then
            If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaProv(vContaFra.NumeroFactura, vContaFra.Anofac)
        
            '---- Poner intconta=1 en ariges.scafac
            b = ActualizarCabFact("tcafpc", cadWHERE, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
'    If Not b Then
'        SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
'        SQL = SQL & " Select *," & DBSet(Mid(cadMen, 1, 200), "T") & " as error From tmpFactu "
'        SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
'        Conn.Execute SQL
'    End If
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaTrans = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaTrans = False
        If Not b Then
            InsertarTMPErrFac cadMen, cadWHERE
'            SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
'            SQL = SQL & " Select *," & DBSet(Mid(cadMen, 1, 200), "T") & " as error From tmpFactu "
'            SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
'            Conn.Execute SQL
        End If
    End If
End Function

Private Function InsertarCabFactTrans(cadWHERE As String, cadErr As String, ByRef Mc As Contadores, FechaFin As String, ByRef vCF As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el a�o de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim sql As String
Dim Sql5 As String
Dim Tipo As Byte
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Nulo4 As String
Dim TipoOpera As Integer
Dim Aux As String
Dim Sql2 As String
Dim ImporAux As Currency

Dim CadenaInsertFaclin2     As String

    On Error GoTo EInsertar
       
    
    sql = " SELECT fecfactu,year(fecrecep) as anofacpr,fecrecep,numfactu,agencias.codmacta,"
    sql = sql & "tcafpc.dtoppago,tcafpc.dtognral,baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    sql = sql & "totalfac,tipoiva1,tipoiva2,tipoiva3, retfacpr, trefacpr, agencias.codtrans, agencias.nomtrans, "
    sql = sql & " tcafpc.domtrans,tcafpc.codpobla,tcafpc.pobtrans,tcafpc.protrans,tcafpc.niftrans,tcafpc.codforpa "
    sql = sql & " FROM " & "tcafpc "
    sql = sql & "INNER JOIN " & "agencias ON tcafpc.codtrans=agencias.codtrans "
    sql = sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
    
        If Mc.ConseguirContador("1", (Rs!FecRecep <= CDate(FechaFin) - 365), True) = 0 Then
            vCF.NumeroFactura = Mc.Contador
            vCF.Anofac = DBLet(Rs!anofacpr)
        
        
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            DtoPPago = Rs!DtoPPago
            DtoGnral = Rs!DtoGnral
            BaseImp = Rs!BaseIVA1 + CCur(DBLet(Rs!BaseIVA2, "N")) + CCur(DBLet(Rs!BaseIVA3, "N"))
            TotalFac = Rs!TotalFac
            AnyoFacPr = Rs!anofacpr
            
            Nulo2 = "N"
            Nulo3 = "N"
            Nulo4 = "N"
            If DBLet(Rs!BaseIVA2, "N") = "0" Then Nulo2 = "S"
            If DBLet(Rs!BaseIVA3, "N") = "0" Then Nulo3 = "S"
            If DBLet(Rs!trefacpr, "N") = "0" Then Nulo4 = "S"
            
            sql = ""
            If vParamAplic.ContabilidadNueva Then sql = "'" & SerieFraPro & "',"
            
            sql = sql & Mc.Contador & "," & DBSet(Rs!FecFactu, "F") & "," & Rs!anofacpr & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!NumFactu, "T") & "," & DBSet(Rs!Codmacta, "T") & "," & ValorNulo & ","
            
            If Not vParamAplic.ContabilidadNueva Then
                sql = sql & DBSet(Rs!BaseIVA1, "N") & "," & DBSet(Rs!BaseIVA2, "N", "S") & "," & DBSet(Rs!BaseIVA3, "N", "S") & ","
                sql = sql & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", Nulo2) & "," & DBSet(Rs!porciva3, "N", Nulo3) & ","
                sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & DBSet(Rs!impoiva2, "N", Nulo2) & "," & DBSet(Rs!impoiva3, "N", Nulo3) & ","
                sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                sql = sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!TipoIVA1, "N") & "," & DBSet(Rs!TipoIVA2, "N", Nulo2) & "," & DBSet(Rs!TipoIVA3, "N", Nulo3) & ",0,"
            Else
            
                sql = sql & DBSet(Rs!NomTrans, "T") & "," & DBSet(Rs!domtrans, "T", "S") & ","
                sql = sql & DBSet(Rs!codPobla, "T", "S") & "," & DBSet(Rs!pobtrans, "T", "S") & "," & DBSet(Rs!protrans, "T", "S") & ","
                sql = sql & DBSet(Rs!NIFTrans, "F", "S") & ",'ES',"
                sql = sql & DBSet(Rs!Codforpa, "N") & ","
                
                TipoOpera = 0
                
                Aux = "0"
                
                'codopera,codconce340,codintra
                sql = sql & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                
                
                'para las lineas
                'factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
                'IVA 1, siempre existe
                Aux = "'" & SerieFraPro & "'," & Mc.Contador & "," & DBSet(Rs!FecRecep, "F") & "," & Rs!anofacpr & ","
                
                Sql2 = Aux & "1," & DBSet(Rs!BaseIVA1, "N") & "," & Rs!TipoIVA1 & "," & DBSet(Rs!porciva1, "N") & ","
                Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
                vTipoIva(0) = Rs!TipoIVA1
                vPorcIva(0) = Rs!porciva1
                vPorcRec(0) = 0
                vImpIva(0) = Rs!impoiva1
                vImpRec(0) = 0
                vBaseIva(0) = Rs!BaseIVA1
                
                vTipoIva(1) = 0: vTipoIva(2) = 0
                
                If Not IsNull(Rs!porciva2) Then
                    Sql2 = Aux & "2," & DBSet(Rs!BaseIVA2, "N") & "," & Rs!TipoIVA2 & "," & DBSet(Rs!porciva2, "N") & ","
                    Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva2, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                    vTipoIva(1) = Rs!TipoIVA2
                    vPorcIva(1) = Rs!porciva2
                    vPorcRec(1) = 0
                    vImpIva(1) = Rs!impoiva2
                    vImpRec(1) = 0
                    vBaseIva(1) = Rs!BaseIVA2
                
                End If
                If Not IsNull(Rs!porciva3) Then
                    Sql2 = Aux & "3," & DBSet(Rs!BaseIVA3, "N") & "," & Rs!TipoIVA3 & "," & DBSet(Rs!porciva3, "N") & ","
                    Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva3, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                    vTipoIva(2) = Rs!TipoIVA3
                    vPorcIva(2) = Rs!porciva3
                    vPorcRec(2) = 0
                    vImpIva(2) = Rs!impoiva3
                    vImpRec(2) = 0
                    vBaseIva(2) = Rs!BaseIVA3
                End If
                
                    
                    
                'Los totales
                'totbases,totbasesret,totivas,totrecargo,totfacpr,
                ImporAux = Rs!BaseIVA1 + DBLet(Rs!BaseIVA2, "N") + DBLet(Rs!BaseIVA3, "N")
                sql = sql & DBSet(ImporAux, "N") & "," & ValorNulo & ","
                'totivas
                ImporAux = Rs!impoiva1 + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
                sql = sql & DBSet(ImporAux, "N") & "," & DBSet(Rs!TotalFac, "N") & ","
                        
                  
'                EsFacturaIntracom2 = ""
'                If DBLet(Rs!tipprove, "N") = 1 Then
'                    'OK es intracomunitaria
'                    EsFacturaIntracom2 = Rs!TipoIVA1
'                End If
                
            End If
            
            
            
            sql = sql & DBSet(Rs!retfacpr, "N", Nulo4) & "," & DBSet(Rs!trefacpr, "N", Nulo4) & ","
            
            If DBSet(Rs!retfacpr, "N", Nulo4) = ValorNulo And DBSet(Rs!trefacpr, "N", Nulo4) = ValorNulo Then
                sql = sql & ValorNulo & ","
            Else
                Sql5 = "select tipo from tcafpc where " & cadWHERE
                Tipo = DevuelveValor(Sql5) ' 0=transportista 1=comisionista
            
                Select Case Tipo
                    Case 0 ' tranportista
                        sql = sql & DBSet(vParamAplic.CtaTraReten, "T") & ","
                    Case 1 ' comisionista
                        sql = sql & DBSet(vParamAplic.CtaComReten, "T") & ","
                End Select
            End If
            
            If vParamAplic.ContabilidadNueva Then sql = sql & "0"
            
            If Not vParamAplic.ContabilidadNueva Then sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
            
            cad = cad & "(" & sql & ")"
            
            If vParamAplic.ContabilidadNueva Then
                sql = "INSERT INTO factpro(numserie,numregis,fecfactu,anofactu,fecharec,fecliqpr,numfactu,codmacta,observa,nommacta,"
                sql = sql & "dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,codforpa,codopera,codconce340,codintra,"
                sql = sql & "totbases,totbasesret,totivas,totfacpr,retfacpr , trefacpr, cuereten, tiporeten)"
            
            
            Else
                'Insertar en la contabilidad
                sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,fecliqpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                sql = sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                sql = sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,nodeducible) "
                sql = sql & " VALUES " & cad
                ConnConta.Execute sql
            End If
            
            If vParamAplic.ContabilidadNueva Then
                'Las  lineas de IVA
                sql = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
                sql = sql & " VALUES " & CadenaInsertFaclin2
                ConnConta.Execute sql
                
            End If
            
            
            'a�adido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            sql = sql & ",'" & DevNombreSQL(Rs!NumFactu) & " @ " & Format(Rs!FecFactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!NomTrans) & "'," & Rs!codTrans & ")"
            conn.Execute sql
            
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactTrans = False
        cadErr = Err.Description
    Else
        InsertarCabFactTrans = True
    End If
End Function

' ### [Monica] 16/01/2008
Public Function InsertarEnTesoreriaNewFac(cadWHERE As String, CtaBan As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim b As Boolean
Dim sql As String, text33csb As String, text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset
Dim rsVenci As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim i As Long
Dim DigConta As String
Dim CC As String
Dim vrefer As String
Dim letraser As String
Dim Rsx As ADODB.Recordset
Dim FecVenci As Date
Dim ImpVenci As Currency
Dim ImpVenci1 As Currency
Dim AcumIva As Currency
Dim PorcIva As String

Dim Rsx7 As ADODB.Recordset
Dim Sql7 As String
Dim CADENA As String

Dim CadRegistro As String
Dim CadRegistro1 As String

' si hay desviacion de importes por redondeo
Dim LineaAModificar As Integer
Dim ImporteACompensar As Currency



    On Error GoTo EInsertarTesoreriaNewFac

    b = False
    InsertarEnTesoreriaNewFac = False
    CadValues = ""
    CadValues2 = ""
    
    Set Rsx = New ADODB.Recordset
    sql = "select * from facturas where " & cadWHERE
    Rsx.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If Not Rsx.EOF Then
    
        '[Monica]22/11/2013: Tema iban
        Sql4 = "select codbanco, codsucur, digcontr, cuentaba, codmacta, ctrolcobroalb, iban, nomclien,domclien,pobclien,codpobla,proclien,cifclien from clientes where codclien = " & DBLet(Rsx!CodClien, "N")
        Set Rs4 = New ADODB.Recordset
        
        Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs4.EOF Then
            letraser = ""
'--monica: 10/02/2009 stipom
'            LetraSer = DevuelveDesdeBDNew(cAgro, "stipom", "letraser", "codtipom", Rsx!codTipoM, "T")
'++monica: 10/02/2009 stipom
            letraser = ObtenerLetraSerie(Rsx!codTipoM)
            
            'insertamos tantos cobros como vtos haya en la forma de pago (hacemos lo que deberia)
            If DBLet(Rs4!ctrolcobroalb) = 0 Or DBLet(Rsx!codTipoM, "T") = "EAC" Then
                
                text33csb = "'Factura:" & DBLet(letraser, "T") & "-" & DBLet(Rsx!NumFactu, "T") & " " & Format(DBLet(Rsx!FecFactu, "F"), "dd/mm/yy") & "'"
                'text41csb = "de " & DBSet(Rsx!TotalFac, "N")
                
'****no teniamos hecho lo del numero de vtos
                'Obtener el N� de Vencimientos de la forma de pago
                sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & DBSet(Rsx!Codforpa, "N")
                Set rsVenci = New ADODB.Recordset
                rsVenci.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                If Not rsVenci.EOF Then
                    If DBLet(rsVenci!numerove, "N") > 0 Then
                        
    '                   textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd,mm,yyyy") & "'"
                
                        CadValuesAux2 = "('" & letraser & "', " & DBSet(Rsx!NumFactu, "N") & ", " & DBSet(Rsx!FecFactu, "F") & ", "
                        '-------- Primer Vencimiento
                        i = 1
                        'FECHA VTO
                        FecVenci = DBLet(Rsx!FecFactu, "F")
                        '=== Laura 23/01/2007
                        'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                        FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                        '===
                        
                        CadValues2 = CadValuesAux2 & i & ", "
                        CadValues2 = CadValues2 & DBSet(Rs4!Codmacta, "T") & ", " & DBSet(Rsx!Codforpa, "N") & ", " & DBSet(FecVenci, "F") & ", "
                        
                        'IMPORTE del Vencimiento
                        If rsVenci!numerove = 1 Then
                            ImpVenci = DBLet(Rsx!TotalFac, "N")
                        Else
                            ImpVenci = Round2(DBLet(Rsx!TotalFac, "N") / rsVenci!numerove, 2)
                            'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                            If ImpVenci * rsVenci!numerove <> DBLet(Rsx!TotalFac, "N") Then
                                ImpVenci = Round2(ImpVenci + (DBLet(Rsx!TotalFac, "N") - ImpVenci * rsVenci!numerove), 2)
                            End If
                        End If
                        
                        CC = DBLet(Rs4!digcontr, "T")
                        If DBLet(Rs4!digcontr, "T") = "**" Then CC = "00"
                        If vParamAplic.ContabilidadNueva Then
                            vvIban = MiFormat(DBLet(Rs4!Iban), "") & MiFormat(DBLet(Rs4!codbanco), "0000") & MiFormat(DBLet(Rs4!codsucur), "0000") & MiFormat(DBLet(CC), "00") & MiFormat(DBLet(Rs4!cuentaba), "0000000000")
                        
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(vvIban, "T", "S") & ", "
                            CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1,"  '),"
                            CadValues2 = CadValues2 & DBSet(Rs4!Nomclien, "T") & "," & DBSet(Rs4!domclien, "T") & "," & DBSet(Rs4!pobclien, "T") & "," & DBSet(Rs4!codPobla, "T") & "," & DBSet(Rs4!proclien, "T") & "," & DBSet(Rs4!cifClien, "T") & ",'ES'"
                        
                            CadValues2 = CadValues2 & "),"
                        
                        Else
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(Rs4!codbanco, "N", "S") & ", " & DBSet(Rs4!codsucur, "N", "S") & ", " & DBSet(CC, "T", "S") & ", " & DBSet(Rs4!cuentaba, "T", "S") & ", "
                            CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1" '),"
                        
                            '[Monica]22/11/2013: tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                CadValues2 = CadValues2 & "," & DBSet(Rs4!Iban, "T", "S") & "),"
                            Else
                                CadValues2 = CadValues2 & "),"
                            End If
                        End If
                    
                        'Resto Vencimientos
                        '--------------------------------------------------------------------
                        For i = 2 To rsVenci!numerove
                           'FECHA Resto Vencimientos
                            '=== Laura 23/01/2007
                            'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                            FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                            '===
                                
                            CadValues2 = CadValues2 & CadValuesAux2 & i & ", " & DBSet(Rs4!Codmacta, "T") & ", " & DBSet(Rsx!Codforpa, "N") & ", '" & Format(FecVenci, FormatoFecha) & "', "
                            
                            'IMPORTE Resto de Vendimientos
                            ImpVenci = Round2(TotalFac / rsVenci!numerove, 2)
                            
                            If vParamAplic.ContabilidadNueva Then
                                vvIban = MiFormat(DBLet(Rs4!Iban), "") & MiFormat(DBLet(Rs4!codbanco), "0000") & MiFormat(DBLet(Rs4!codsucur), "0000") & MiFormat(DBLet(CC), "00") & MiFormat(DBLet(Rs4!cuentaba), "0000000000")
                                
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(vvIban, "T", "S") & ", "
                                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1,"  '),"
                                CadValues2 = CadValues2 & DBSet(Rs4!Nomclien, "T") & "," & DBSet(Rs4!domclien, "T") & "," & DBSet(Rs4!pobclien, "T") & "," & DBSet(Rs4!codPobla, "T") & "," & DBSet(Rs4!proclien, "T") & "," & DBSet(Rs4!cifClien, "T") & ",'ES'"
                                CadValues2 = CadValues2 & "),"
                            Else
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(Rs4!codbanco, "N", "S") & ", " & DBSet(Rs4!codsucur, "N", "S") & ", " & DBSet(CC, "T", "S") & ", " & DBSet(Rs4!cuentaba, "T", "S") & ", "
                                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1" '),"
                            
                                '[Monica]22/11/2013: tema iban
                                If vEmpresa.HayNorma19_34Nueva = 1 Then
                                    CadValues2 = CadValues2 & "," & DBSet(Rs4!Iban, "T", "S") & "),"
                                Else
                                    CadValues2 = CadValues2 & "),"
                                End If
                            End If
                        Next i
                        ' quitamos la ultima coma
                        CadValues2 = Mid(CadValues2, 1, Len(CadValues2) - 1)
                            
                        If vParamAplic.ContabilidadNueva Then
                            'Insertamos en la tabla scobro de la CONTA
                            sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                            sql = sql & " ctabanc1, iban, fecultco, impcobro, "
                            sql = sql & " text33csb, text41csb, agente, "
                            sql = sql & " nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais) "
                        Else
                            'Insertamos en la tabla scobro de la CONTA
                            sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                            sql = sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                            sql = sql & " text33csb, text41csb, agente"
                            '[Monica]22/11/2013: Tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                sql = sql & ", iban) "
                            Else
                                sql = sql & ") "
                            End If
                        End If
                        sql = sql & " VALUES " & CadValues2
                        ConnConta.Execute sql
                    
                    End If
                End If
    
'****
'                CadValuesAux2 = "(" & DBSet(LetraSer, "T") & "," & DBSet(Rsx!numfactu, "N") & "," & DBSet(Rsx!fecfactu, "F") & ", 1," & DBSet(Rs4!Codmacta, "T") & ","
'                CadValues2 = CadValuesAux2 & DBSet(Rsx!Codforpa, "N") & "," & DBSet(Rsx!fecfactu, "F") & "," & DBSet(Rsx!TotalFac, "N") & ","
'                CadValues2 = CadValues2 & DBSet(CtaBan, "T") & "," & DBSet(Rs4!codbanco, "N") & "," & DBSet(Rs4!codsucur, "N") & ","
'                CadValues2 = CadValues2 & DBSet(CC, "T") & "," & DBSet(Rs4!cuentaba, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
'                CadValues2 = CadValues2 & text33csb & "," & DBSet(text41csb, "T") & ",1)"
                
    
            Else
                ' cliente.ctrolcobroalb = 1
                ' insertamos: un registro por linea de iva de cada variedad y un registro por cada linea de iva de envase (antes por el total de iva)
                '             un registro por el total de envases
                '             un registro por cada linea de factura_variedad (que son las lineas de albaran facturadas)
                
                text33csb = "'Factura:" & DBLet(letraser, "T") & "-" & DBLet(Rsx!NumFactu, "T") & " " & Format(DBLet(Rsx!FecFactu, "F"), "dd/mm/yy") & "'"
                'text41csb = "de " & DBSet(Rsx!TotalFac, "N")
            
                CadValuesAux2 = "('" & letraser & "', " & DBSet(Rsx!NumFactu, "N") & ", " & DBSet(Rsx!FecFactu, "F") & ", "
                CadValues2 = ""
                
                CC = DBLet(Rs4!digcontr, "T")
                If DBLet(Rs4!digcontr, "T") = "**" Then CC = "00"
                
                FecVenci = DBLet(Rsx!FecFactu, "F")
'[Monica]27/09/2010: la fecha de vencimiento tiene que ser la de factura pero sumandole los dias del primer vencimiento
                sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & DBSet(Rsx!Codforpa, "N")
                Set rsVenci = New ADODB.Recordset
                rsVenci.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not rsVenci.EOF Then
                    FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                End If
                Set rsVenci = Nothing
'[Monica]27/09/2010:end

                '-------- Primer Vencimiento ---> IVA
                ImpVenci = DBLet(Rsx!impoiva1, "N") + DBLet(Rsx!impoiva2, "N") + DBLet(Rsx!impoiva3, "N") + DBLet(Rsx!imporec1, "N") + DBLet(Rsx!imporec2, "N") + DBLet(Rsx!imporec3, "N")
                i = 0
'[Monica] 01/04/2010 : antes teniamos un solo registro por el iva total (ahora tenemos que cuadrarlo)
'                If ImpVenci <> 0 Then
'                    I = I + 1
'                    CadValues2 = CadValues2 & CadValuesAux2 & I & ", "
'                    CadValues2 = CadValues2 & DBSet(Rs4!Codmacta, "T") & ", " & DBSet(Rsx!Codforpa, "N") & ", " & DBSet(FecVenci, "F") & ", "
'
'                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(Rs4!codbanco, "N") & ", " & DBSet(Rs4!codsucur, "N") & ", " & DBSet(CC, "T") & ", " & DBSet(Rs4!cuentaba, "T") & ", "
'                    CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1,'IVA'),"
'                End If
'[Monica] 01/04/2010 : sustituido por un vencimiento por cada iva de linea de variedad
                ' y otra linea por cada iva de linea de envase
                Sql7 = "select 0 tipo, numalbar, numlinealbar numlinea, impornet importe, codigiva from facturas_variedad where " & Replace(cadWHERE, "facturas", "facturas_variedad")
                Sql7 = Sql7 & " union "
                '[Monica]11/02/2013: quieren en la referencia del iva el numero de albaran que pongan
                If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                    Sql7 = Sql7 & " select 1 tipo, numalbar numalbar, 0 numlinea, importel importe, codigiva from facturas_envases where " & Replace(cadWHERE, "facturas", "facturas_envases")
                Else
                    Sql7 = Sql7 & " select 1 tipo,0, numlinea numlinea, importel importe, codigiva from facturas_envases where " & Replace(cadWHERE, "facturas", "facturas_envases")
                End If
                Sql7 = Sql7 & " order by 1 "
                Set Rsx7 = New ADODB.Recordset
                Rsx7.Open Sql7, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                AcumIva = 0
                ImporteACompensar = 0
                LineaAModificar = 0
                
                CadRegistro = ""
                CadRegistro1 = ""
                
                While Not Rsx7.EOF
                    i = i + 1
                    
                    PorcIva = ""
                    PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", DBLet(Rsx7!Codigiva, "N"), "N")
                    If PorcIva = "" Then PorcIva = "0"
                    
                    ImpVenci1 = Round2(DBLet(Rsx7!Importe, "N") * CCur(PorcIva) / 100, 2)
                    
                    AcumIva = AcumIva + ImpVenci1
                    
                    CadValues2 = CadValues2 & CadValuesAux2 & i & ","
                    CadValues2 = CadValues2 & DBSet(Rs4!Codmacta, "T") & ", " & DBSet(Rsx!Codforpa, "N") & ", " & DBSet(FecVenci, "F") & ", "
                    
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(DBLet(Rs4!Iban), "") & MiFormat(DBLet(Rs4!codbanco), "0000") & MiFormat(DBLet(Rs4!codsucur), "0000") & MiFormat(CC, "00") & MiFormat(DBLet(Rs4!cuentaba), "0000000000")
                        
                        CadValues2 = CadValues2 & DBSet(ImpVenci1, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(vvIban, "T") & ", "
                        CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1,"
                    Else
                        CadValues2 = CadValues2 & DBSet(ImpVenci1, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(Rs4!codbanco, "N", "S") & ", " & DBSet(Rs4!codsucur, "N", "S") & ", " & DBSet(CC, "T", "S") & ", " & DBSet(Rs4!cuentaba, "T", "S") & ", "
                        CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1,"
                    End If
                        
                    If DBLet(Rsx7!Tipo, "N") = 0 Then
                        CadValues2 = CadValues2 & "'IVA VARIEDAD'," & DBSet(Format(DBLet(Rsx7!NumAlbar, "N"), "0000000"), "T") & "," & DBSet(Format(DBLet(Rsx7!numlinea, "N"), "000"), "T")
                    Else
                        '[Monica]11/02/2013: metemos en la referencia el nro de albaran que hayan metido
                        If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                            CadValues2 = CadValues2 & "'IVA ENVASE'," & DBSet(DBLet(Rsx7!NumAlbar, "N"), "T") & "," & ValorNulo
                        Else
                            CadValues2 = CadValues2 & "'IVA ENVASE'," & DBSet(Format(DBLet(Rsx7!numlinea, "N"), "000"), "T") & "," & ValorNulo
                        End If
                    End If
                    
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & "," & DBSet(Rs4!Nomclien, "T") & "," & DBSet(Rs4!domclien, "T") & "," & DBSet(Rs4!pobclien, "T") & "," & DBSet(Rs4!codPobla, "T") & "," & DBSet(Rs4!proclien, "T") & "," & DBSet(Rs4!cifClien, "T") & ",'ES'"

                        CadValues2 = CadValues2 & "),"
                    Else
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & ", " & DBSet(Rs4!Iban, "T", "S") & "),"
                        Else
                            CadValues2 = CadValues2 & "),"
                        End If
                    End If
                    Rsx7.MoveNext
                Wend
                Set Rsx7 = Nothing
                    
                If AcumIva <> ImpVenci Then
                    LineaAModificar = i
                    ImporteACompensar = ImpVenci - AcumIva
                End If
                
                
                
                '-------- Segundo Vencimiento ---> TOTAL de ENVASES (si no es Picassent 07/02/2013)
                '                                  En caso de ser Picassent, es un cobro por envase.
                If vParamAplic.Cooperativa <> 2 Then
                    Sql7 = "select sum(importel) from facturas_envases where " & Replace(cadWHERE, "facturas", "facturas_envases")
                    Set Rsx7 = New ADODB.Recordset
                    
                    Rsx7.Open Sql7, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    If Not Rsx7.EOF Then
                        If DBLet(Rsx7.Fields(0).Value, "N") <> 0 Then
                            i = i + 1
                            
                            CadValues2 = CadValues2 & CadValuesAux2 & i & ", "
                            CadValues2 = CadValues2 & DBSet(Rs4!Codmacta, "T") & ", " & DBSet(Rsx!Codforpa, "N") & ", " & DBSet(FecVenci, "F") & ", "
                                    
                            
                            ImpVenci = DBLet(Rsx7.Fields(0).Value, "N")
                            
                            If vParamAplic.ContabilidadNueva Then
                                vvIban = MiFormat(DBLet(Rs4!Iban), "") & MiFormat(DBLet(Rs4!codbanco), "0000") & MiFormat(DBLet(Rs4!codsucur), "0000") & MiFormat(CC, "00") & MiFormat(DBLet(Rs4!cuentaba), "0000000000")
                            
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(vvIban, "T", "S") & ", "
                                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1,'ENVASES'," & ValorNulo & "," & ValorNulo & ","  '& "),"
                                CadValues2 = CadValues2 & DBSet(Rs4!Nomclien, "T") & "," & DBSet(Rs4!domclien, "T") & "," & DBSet(Rs4!pobclien, "T") & "," & DBSet(Rs4!codPobla, "T") & "," & DBSet(Rs4!proclien, "T") & "," & DBSet(Rs4!cifClien, "T") & ",'ES'"
                                CadValues2 = CadValues2 & "),"
                            
                            Else
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(Rs4!codbanco, "N", "S") & ", " & DBSet(Rs4!codsucur, "N", "S") & ", " & DBSet(CC, "T", "S") & ", " & DBSet(Rs4!cuentaba, "T", "S") & ", "
                                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1,'ENVASES'," & ValorNulo & "," & ValorNulo '& "),"
                                
                                '[Monica]22/11/2013: Tema iban
                                If vEmpresa.HayNorma19_34Nueva = 1 Then
                                    CadValues2 = CadValues2 & ", " & DBSet(Rs4!Iban, "T", "S") & "),"
                                Else
                                    CadValues2 = CadValues2 & "),"
                                End If
                            End If
                        End If
                    End If
                Else
                    Sql7 = "select numlinea, numalbar, importel from facturas_envases where " & Replace(cadWHERE, "facturas", "facturas_envases")
                    Set Rsx7 = New ADODB.Recordset
                    
                    Rsx7.Open Sql7, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    If Not Rsx7.EOF Then Rsx7.MoveFirst
                    
                    While Not Rsx7.EOF
                        If DBLet(Rsx7.Fields(2).Value, "N") <> 0 Then
                            i = i + 1
                            
                            CadValues2 = CadValues2 & CadValuesAux2 & i & ", "
                            CadValues2 = CadValues2 & DBSet(Rs4!Codmacta, "T") & ", " & DBSet(Rsx!Codforpa, "N") & ", " & DBSet(FecVenci, "F") & ", "
                                    
                            ImpVenci = DBLet(Rsx7.Fields(2).Value, "N")
                                    
                            If vParamAplic.ContabilidadNueva Then
                                vvIban = MiFormat(DBLet(Rs4!Iban), "") & MiFormat(DBLet(Rs4!codbanco), "0000") & MiFormat(DBLet(Rs4!codsucur), "0000") & MiFormat(CC, "00") & MiFormat(DBLet(Rs4!cuentaba), "0000000000")
                                
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(vvIban, "T", "S") & ", "
                                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1,'ENVASES'," & DBSet(Rsx7.Fields(1).Value, "T") & "," & ValorNulo  '& "),"
                                CadValues2 = CadValues2 & DBSet(Rs4!Nomclien, "T") & "," & DBSet(Rs4!domclien, "T") & "," & DBSet(Rs4!pobclien, "T") & "," & DBSet(Rs4!codPobla, "T") & "," & DBSet(Rs4!proclien, "T") & "," & DBSet(Rs4!cifClien, "T") & ",'ES'"
                                CadValues2 = CadValues2 & "),"
                            Else
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(Rs4!codbanco, "N", "S") & ", " & DBSet(Rs4!codsucur, "N", "S") & ", " & DBSet(CC, "T", "S") & ", " & DBSet(Rs4!cuentaba, "T", "S") & ", "
                                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1,'ENVASES'," & DBSet(Rsx7.Fields(1).Value, "T") & "," & ValorNulo '& "),"
                                
                                '[Monica]22/11/2013: Tema iban
                                If vEmpresa.HayNorma19_34Nueva = 1 Then
                                    CadValues2 = CadValues2 & ", " & DBSet(Rs4!Iban, "T", "S") & "),"
                                Else
                                    CadValues2 = CadValues2 & "),"
                                End If
                                
                            End If
                        End If
                        
                        Rsx7.MoveNext
                    Wend
                
                End If
                
                Rsx7.Close
                Set Rsx7 = Nothing
            
                '-------- Resto de Vencimientos ---> uno por cada albaran linea
                Sql7 = "select * from facturas_variedad where " & Replace(cadWHERE, "facturas", "facturas_variedad")
                Set Rsx7 = New ADODB.Recordset
                
                Rsx7.Open Sql7, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not Rsx7.EOF
                    If DBLet(Rsx7!impornet, "N") <> 0 Then
                        i = i + 1
                        
                        CadValues2 = CadValues2 & CadValuesAux2 & i & ", "
                        CadValues2 = CadValues2 & DBSet(Rs4!Codmacta, "T") & ", " & DBSet(Rsx!Codforpa, "N") & ", " & DBSet(FecVenci, "F") & ", "
                                
                        ImpVenci = DBLet(Rsx7!impornet, "N")
                                
                        If vParamAplic.ContabilidadNueva Then
                            vvIban = MiFormat(DBLet(Rs4!Iban), "") & MiFormat(DBLet(Rs4!codbanco), "0000") & MiFormat(DBLet(Rs4!codsucur), "0000") & MiFormat(CC, "00") & MiFormat(DBLet(Rs4!cuentaba), "0000000000")
                            
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(vvIban, "T", "S") & ", "
                            CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1,"
                        Else
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(Rs4!codbanco, "N", "S") & ", " & DBSet(Rs4!codsucur, "N", "S") & ", " & DBSet(CC, "T", "S") & ", " & DBSet(Rs4!cuentaba, "T", "S") & ", "
                            CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1,"
                        End If
                        
                        ' [Monica]13/01/2011: por Picassent si se lo indicamos introducimos la referencia de linea de albaran
                        If vParamAplic.PaseRefLineaAlb Then
                            If DBLet(Trim(DevuelveDesdeBDNew(cAgro, "albaran_variedad", "referencia", "numalbar", Rsx7!NumAlbar, "T")), "T") <> "" Then ' si tiene valor la referencia de linea
                                ' referencia, referencia1, referencia2
                                CADENA = DBSet(Trim(DBLet(DevuelveDesdeBDNew(cAgro, "albaran_variedad", "referencia", "numalbar", Rsx7!NumAlbar, "T"), "T")), "T") & "," & DBSet(Format(DBLet(Rsx7!NumAlbar, "N"), "0000000"), "T") & "," & DBSet(Format(DBLet(Rsx7!numlinealbar, "N"), "000"), "T")
                            Else ' si no la referencia de la cabecera
                                CADENA = DBSet(Trim(DevuelveDesdeBDNew(cAgro, "albaran", "refclien", "numalbar", Rsx7!NumAlbar, "T")), "T") & "," & DBSet(Format(DBLet(Rsx7!NumAlbar, "N"), "0000000"), "T") & "," & DBSet(Format(DBLet(Rsx7!numlinealbar, "N"), "000"), "T")
                            End If
                        Else
                        ' [Monica] 13/01/2010 hasta aqui
                            ' referencia, referencia1, referencia2
                            CADENA = DBSet(Trim(DevuelveDesdeBDNew(cAgro, "albaran", "refclien", "numalbar", Rsx7!NumAlbar, "T")), "T") & "," & DBSet(Format(DBLet(Rsx7!NumAlbar, "N"), "0000000"), "T") & "," & DBSet(Format(DBLet(Rsx7!numlinealbar, "N"), "000"), "T")
                        End If
                        
                        CadValues2 = CadValues2 & CADENA '& "),"
                        
                        If vParamAplic.ContabilidadNueva Then
                            CadValues2 = CadValues2 & "," & DBSet(Rs4!Nomclien, "T") & "," & DBSet(Rs4!domclien, "T") & "," & DBSet(Rs4!pobclien, "T") & "," & DBSet(Rs4!codPobla, "T") & "," & DBSet(Rs4!proclien, "T") & "," & DBSet(Rs4!cifClien, "T") & ",'ES'"
                            CadValues2 = CadValues2 & "),"
                        Else
                            '[Monica]22/11/2013: Tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                CadValues2 = CadValues2 & ", " & DBSet(Rs4!Iban, "T", "S") & "),"
                            Else
                                CadValues2 = CadValues2 & "),"
                            End If
                        End If
                    End If
                    Rsx7.MoveNext
                Wend
            
                Set Rsx7 = Nothing
            
                If i > 0 Then
                    'quitamos la ultima coma
                    CadValues2 = Mid(CadValues2, 1, Len(CadValues2) - 1)
                
                
                    If vParamAplic.ContabilidadNueva Then
                        'Insertamos en la tabla scobro de la CONTA
                        sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        sql = sql & "ctabanc1, iban, fecultco, impcobro, "
                        sql = sql & " text33csb, text41csb, agente, referencia, referencia1, referencia2," ') "
                        sql = sql & "nomclien,domclien,pobclien,cpclien,proclien,nifclien,codpais"
                        sql = sql & ") "
                    Else
                        'Insertamos en la tabla scobro de la CONTA
                        sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        sql = sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                        sql = sql & " text33csb, text41csb, agente, referencia, referencia1, referencia2" ') "
                        
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            sql = sql & ", iban) "
                        Else
                            sql = sql & ") "
                        End If
                    End If
                    sql = sql & " VALUES " & CadValues2
                    ConnConta.Execute sql
                    
                    
                    If ImporteACompensar <> 0 Then
                        If vParamAplic.ContabilidadNueva Then
                            sql = "update cobros set impvenci = impvenci + " & DBSet(ImporteACompensar, "N")
                            sql = sql & " where numserie = " & DBSet(letraser, "T")
                            sql = sql & " and numfactu = " & DBSet(Rsx!NumFactu, "N")
                            sql = sql & " and fecfactu = " & DBSet(Rsx!FecFactu, "F")
                            sql = sql & " and numorden = " & DBSet(LineaAModificar, "N")
                            
                            ConnConta.Execute sql
                        
                        Else
                            sql = "update scobro set impvenci = impvenci + " & DBSet(ImporteACompensar, "N")
                            sql = sql & " where numserie = " & DBSet(letraser, "T")
                            sql = sql & " and codfaccl = " & DBSet(Rsx!NumFactu, "N")
                            sql = sql & " and fecfaccl = " & DBSet(Rsx!FecFactu, "F")
                            sql = sql & " and numorden = " & DBSet(LineaAModificar, "N")
                            
                            ConnConta.Execute sql
                        End If
                    End If
                End If
            End If
        End If
    
        b = True
    End If
    
EInsertarTesoreriaNewFac:
    If Err.Number <> 0 Then
        b = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaNewFac = b
End Function


' ### [Monica] Insertar en tesoreria las facturas de venta de socio
Public Function InsertarEnTesoreriaNewFacSoc(cadWHERE As String, CtaBan As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim b As Boolean
Dim sql As String, text33csb As String, text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset
Dim rsVenci As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim i As Byte
Dim DigConta As String
Dim CC As String
Dim vrefer As String
Dim letraser As String
Dim Rsx As ADODB.Recordset
Dim FecVenci As Date
Dim ImpVenci As Currency
Dim ImpVenci1 As Currency
Dim AcumIva As Currency
Dim PorcIva As String

Dim Rsx7 As ADODB.Recordset
Dim Sql7 As String
Dim CADENA As String

Dim CadRegistro As String
Dim CadRegistro1 As String

' si hay desviacion de importes por redondeo
Dim LineaAModificar As Integer
Dim ImporteACompensar As Currency

Dim SeccionHorto As Integer

    On Error GoTo EInsertarTesoreriaNewFacSoc

    b = False
    InsertarEnTesoreriaNewFacSoc = False
    CadValues = ""
    CadValues2 = ""
    
    Set Rsx = New ADODB.Recordset
    sql = "select * from facturassocio where " & cadWHERE
    Rsx.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If Not Rsx.EOF Then
        SeccionHorto = DevuelveValor("select seccionhorto from rparam")
        '[Monica]22/11/2013: Tema Iban
        Sql4 = "select codbanco, codsucur, digcontr, cuentaba, rsocios_seccion.codmaccli codmacta, iban, rsocios.nomsocio,rsocios.dirsocio,rsocios.pobsocio,rsocios.prosocio,rsocios.codpostal,rsocios.nifsocio from rsocios, rsocios_seccion where rsocios.codsocio = " & DBLet(Rsx!CodSocio, "N")
        Sql4 = Sql4 & " and rsocios_seccion.codsocio = rsocios.codsocio and rsocios_seccion.codsecci = " & DBSet(SeccionHorto, "N")
        Set Rs4 = New ADODB.Recordset
        
        Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs4.EOF Then
            letraser = ""
            letraser = ObtenerLetraSerie(Rsx!codTipoM)
            
            text33csb = "'Factura:" & DBLet(letraser, "T") & "-" & DBLet(Rsx!NumFactu, "T") & " " & Format(DBLet(Rsx!FecFactu, "F"), "dd/mm/yy") & "'"
            'text41csb = "de " & DBSet(Rsx!TotalFac, "N")
            
            'Obtener el N� de Vencimientos de la forma de pago
            sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & DBSet(Rsx!Codforpa, "N")
            Set rsVenci = New ADODB.Recordset
            rsVenci.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

            If Not rsVenci.EOF Then
                If DBLet(rsVenci!numerove, "N") > 0 Then
                    
'                   textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd,mm,yyyy") & "'"
            
                    CadValuesAux2 = "('" & letraser & "', " & DBSet(Rsx!NumFactu, "N") & ", " & DBSet(Rsx!FecFactu, "F") & ", "
                    '-------- Primer Vencimiento
                    i = 1
                    'FECHA VTO
                    FecVenci = DBLet(Rsx!FecFactu, "F")
                    '=== Laura 23/01/2007
                    'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                    FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                    '===
                    
                    CadValues2 = CadValuesAux2 & i & ", "
                    CadValues2 = CadValues2 & DBSet(Rs4!Codmacta, "T") & ", " & DBSet(Rsx!Codforpa, "N") & ", " & DBSet(FecVenci, "F") & ", "
                    
                    'IMPORTE del Vencimiento
                    If rsVenci!numerove = 1 Then
                        ImpVenci = DBLet(Rsx!TotalFac, "N")
                    Else
                        ImpVenci = Round2(DBLet(Rsx!TotalFac, "N") / rsVenci!numerove, 2)
                        'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                        If ImpVenci * rsVenci!numerove <> DBLet(Rsx!TotalFac, "N") Then
                            ImpVenci = Round2(ImpVenci + (DBLet(Rsx!TotalFac, "N") - ImpVenci * rsVenci!numerove), 2)
                        End If
                    End If
                    
                    CC = DBLet(Rs4!digcontr, "T")
                    If DBLet(Rs4!digcontr, "T") = "**" Then CC = "00"
                    
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(DBLet(Rs4!Iban), "") & MiFormat(DBLet(Rs4!codbanco), "0000") & MiFormat(DBLet(Rs4!codsucur), "0000") & MiFormat(CC, "00") & MiFormat(DBLet(Rs4!cuentaba), "0000000000")
                    
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(vvIban, "T", "S") & ","
                        CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1,"
                        CadValues2 = CadValues2 & DBSet(Rs4!NomSocio, "T") & "," & DBSet(Rs4!dirsocio, "T") & "," & DBSet(Rs4!pobsocio, "T") & "," & DBSet(Rs4!codPostal, "T") & "," & DBSet(Rs4!prosocio, "T") & "," & DBSet(Rs4!nifsocio, "T") & ",'ES'"
                        CadValues2 = CadValues2 & "),"
                    
                    Else
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(Rs4!codbanco, "N", "S") & ", " & DBSet(Rs4!codsucur, "N", "S") & ", " & DBSet(CC, "T", "S") & ", " & DBSet(Rs4!cuentaba, "T", "S") & ", "
                        CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1" '),"
                        
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & ", " & DBSet(Rs4!Iban, "T", "S") & "),"
                        Else
                            CadValues2 = CadValues2 & "),"
                        End If
                    End If
                
                    'Resto Vencimientos
                    '--------------------------------------------------------------------
                    For i = 2 To rsVenci!numerove
                       'FECHA Resto Vencimientos
                        '=== Laura 23/01/2007
                        'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                        FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                        '===
                            
                        CadValues2 = CadValues2 & CadValuesAux2 & i & ", " & DBSet(Rs4!Codmacta, "T") & ", " & DBSet(Rsx!Codforpa, "N") & ", '" & Format(FecVenci, FormatoFecha) & "', "
                        
                        'IMPORTE Resto de Vendimientos
                        ImpVenci = Round2(TotalFac / rsVenci!numerove, 2)
                        
                        If vParamAplic.ContabilidadNueva Then
                            vvIban = MiFormat(DBLet(Rs4!Iban), "") & MiFormat(DBLet(Rs4!codbanco), "0000") & MiFormat(DBLet(Rs4!codsucur), "0000") & MiFormat(CC, "00") & MiFormat(DBLet(Rs4!cuentaba), "0000000000")
                        
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(vvIban, "T", "S") & ","
                            CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1,"
                            CadValues2 = CadValues2 & DBSet(Rs4!NomSocio, "T") & "," & DBSet(Rs4!dirsocio, "T") & "," & DBSet(Rs4!pobsocio, "T") & "," & DBSet(Rs4!codPostal, "T") & "," & DBSet(Rs4!prosocio, "T") & "," & DBSet(Rs4!nifsocio, "T") & ",'ES'"
                            CadValues2 = CadValues2 & "),"
                        
                        Else
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(Rs4!codbanco, "N", "S") & ", " & DBSet(Rs4!codsucur, "N", "S") & ", " & DBSet(CC, "T", "S") & ", " & DBSet(Rs4!cuentaba, "T", "S") & ", "
                            CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1" '),"
                            
                            '[Monica]22/11/2013: Tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                CadValues2 = CadValues2 & ", " & DBSet(Rs4!Iban, "T", "S") & "),"
                            Else
                                CadValues2 = CadValues2 & "),"
                            End If
                        End If
                    Next i
                    ' quitamos la ultima coma
                    CadValues2 = Mid(CadValues2, 1, Len(CadValues2) - 1)
                        
                    'Insertamos en la tabla scobro de la CONTA
                    If vParamAplic.ContabilidadNueva Then
                        sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        sql = sql & "ctabanc1, iban, fecultco, impcobro, "
                        sql = sql & " text33csb, text41csb, agente,nomclien,domclien,pobclien,cpclien,proclien,nifclien,codpais) " ') "
                    
                    Else
                        sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        sql = sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                        sql = sql & " text33csb, text41csb, agente" ') "
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            sql = sql & ", iban) "
                        Else
                            sql = sql & ") "
                        End If
                    End If
                    
                    sql = sql & " VALUES " & CadValues2
                    ConnConta.Execute sql
                End If
            End If
        End If
    
        b = True
    End If
    
EInsertarTesoreriaNewFacSoc:
    If Err.Number <> 0 Then
        b = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaNewFacSoc = b
End Function



Private Function HayFacturasACuenta() As Boolean
Dim sql As String

    sql = "select count(*) from tmpfactu where codtipom = 'EAC'"
    
    HayFacturasACuenta = (TotalRegistros(sql) <> 0)

End Function


Public Function InsertarAsientoDiario(FecEnt As String, NDiario As String, CtaContra As String, NLiq As String, FecLiq As String, cadErr As String)
Dim sql As String
Dim numdocum As String
Dim Ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim i As Long
Dim b As Boolean
Dim cad As String
Dim FeFact As Date
Dim cadMen As String

Dim letraser As String
Dim Concep As Integer
Dim Amplia As String
Dim tipoF As String
Dim Conceptoh As String
Dim Conceptod As String
Dim Rs As ADODB.Recordset
Dim Mc As Contadores
Dim Total As Currency

    conn.BeginTrans
    ConnConta.BeginTrans


    Screen.MousePointer = vbHourglass

    Set Mc = New Contadores
    If Mc.ConseguirContador("0", (FecEnt <= CDate(vEmpresa.FechaFin)), True) = 0 Then

        Obs = "Contab.Pago Anecoop Liquidaci�n " & NLiq & " de Fecha " & Format(FecLiq, "dd/mm/yyyy")
    
        'Insertar en la conta Cabecera Asiento
        b = InsertarCabAsientoDia(NDiario, Mc.Contador, CStr(Format(FecEnt, "dd/mm/yyyy")), Obs, cadMen)
        cadMen = "Insertando Cab. Asiento: " & cadMen
    
        If vParamAplic.ContabilidadNueva Then
        
            sql = "select distinct * from tmpinformes, ariconta" & vParamAplic.NumeroConta & ".cobros cc "
            sql = sql & " where codusu = " & vUsu.Codigo
            sql = sql & " and tmpinformes.nombre1 = cc.numserie "
            sql = sql & " and tmpinformes.importe1 = cc.numfactu "
            sql = sql & " and tmpinformes.fecha1 = cc.fecfactu "
            sql = sql & " and tmpinformes.importe2 = cc.numorden "
            sql = sql & " order by importe1, fecha1, importe2 "
        
        Else
    
            sql = "select distinct * from tmpinformes, conta" & vParamAplic.NumeroConta & ".scobro cc "
            sql = sql & " where codusu = " & vUsu.Codigo
            sql = sql & " and tmpinformes.nombre1 = cc.numserie "
            sql = sql & " and tmpinformes.importe1 = cc.codfaccl "
            sql = sql & " and tmpinformes.fecha1 = cc.fecfaccl "
            sql = sql & " and tmpinformes.importe2 = cc.numorden "
            sql = sql & " order by importe1, fecha1, importe2 "
            
        End If
        Set Rs = New ADODB.Recordset
        Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        i = 0
        ImporteD = 0
        ImporteH = 0
    
        b = True
        
        While Not Rs.EOF And b
            If vParamAplic.ContabilidadNueva Then
                numdocum = letraser & Format(Rs!NumFactu, "0000000")
                tipoF = DevuelveValor("select tipforpa from formapago where codforpa = " & DBSet(Rs!Codforpa, "N"))
                
                Conceptoh = "conhacli"
                Conceptod = DevuelveDesdeBDNew(cConta, "tipofpago", "condecli", "tipoformapago", tipoF, "N", Conceptoh)
            Else
                numdocum = letraser & Format(Rs!Codfaccl, "0000000")
                tipoF = DevuelveValor("select tipoforp from forpago where codforpa = " & DBSet(Rs!Codforpa, "N"))
            
                Conceptoh = "conhacli"
                Conceptod = DevuelveDesdeBDNew(cConta, "stipoformapago", "condecli", "tipoformapago", tipoF, "N", Conceptoh)
            End If
            
            
            Amplia = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", CStr(Conceptod), "N"))
            ampliaciond = Amplia & " " & Format(Rs!referencia1, "0000000") & "-" & DBLet(Rs!referencia2)
            
            Amplia = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", CStr(Conceptoh), "N"))
            ampliacionh = Amplia & " " & Format(Rs!referencia1, "0000000") & "-" & DBLet(Rs!referencia2)
                
            i = i + 1
            
            cad = DBSet(NDiario, "N") & "," & DBSet(FecEnt, "F") & "," & DBSet(Mc.Contador, "N") & ","
            cad = cad & DBSet(i, "N") & "," & DBSet(Rs!Codmacta, "T") & "," & DBSet(numdocum, "T") & ","
            
            ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
            If DBLet(Rs!ImpVenci, "N") > 0 Then
                ' importe al haber en positivo
                cad = cad & DBSet(Conceptoh, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                cad = cad & DBSet(Rs!ImpVenci, "N") & "," & ValorNulo & "," & DBSet(CtaContra, "T") & "," & ValorNulo & ",0"
            
                ImporteH = ImporteH + CCur(DBLet(Rs!ImpVenci, "N"))
            Else
                ' importe al debe en positivo cambiamos signo
                cad = cad & DBSet(Conceptod, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(DBLet(Rs!ImpVenci, "N") * (-1), "N") & ","
                cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaContra, "T") & "," & ValorNulo & ",0"
            
                ImporteD = ImporteD + CCur(DBLet(Rs!ImpVenci, "N") * (-1))
            End If
            
            cad = "(" & cad & ")"
            
            b = InsertarLinAsientoDia(cad, cadMen)
            cadMen = "Insertando Lin. Asiento: " & i
            
            
            Rs.MoveNext
            
        Wend
    
        If b Then
    
            i = i + 1
                    
            numdocum = Format(NLiq, "0000000")
                    
            ' el Total es sobre la cuenta del cliente
            cad = DBSet(NDiario, "N") & "," & DBSet(FecEnt, "F") & "," & DBSet(Mc.Contador, "N") & ","
            cad = cad & DBSet(i, "N") & ","
            cad = cad & DBSet(CtaContra, "T") & "," & DBSet(numdocum, "T") & ","
                
            Total = ImporteH - ImporteD
                
            ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
            If Total > 0 Then
                ' importe al debe en positivo
                cad = cad & DBSet(Conceptod, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Total, "N") & ","
                cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
            Else
                ' importe al haber en positivo, cambiamos el signo
                cad = cad & DBSet(Conceptoh, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                cad = cad & DBSet(Total * (-1), "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
            End If
            
            cad = "(" & cad & ")"
            
            b = InsertarLinAsientoDia(cad, cadMen)
            cadMen = "Insertando Lin. Asiento: " & i
        End If
        
        If b Then b = EliminarCobros(cadMen)
        If b Then b = MarcarRegistros(cadMen)
        
        Set Mc = Nothing
        
        Rs.Close
        Set Rs = Nothing
    End If
    
EInsertar:
    
    Screen.MousePointer = vbDefault
    
    
    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
        ConnConta.RollbackTrans
        InsertarAsientoDiario = False
        cadErr = Err.Description
    Else
        conn.CommitTrans
        ConnConta.CommitTrans
        InsertarAsientoDiario = True
    End If
End Function






Public Function InsertarCabAsientoDia(Diario As String, Asiento As String, fecha As String, Obs As String, cadErr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String

    On Error GoTo EInsertar
       
    
    If vParamAplic.ContabilidadNueva Then
        cad = Format(Diario, "00") & ", " & DBSet(fecha, "F") & "," & Format(Asiento, "000000") & ","
        cad = cad & DBSet(Obs, "T") & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARIAGRO COMERCIAL'"
        
        cad = "(" & cad & ")"
    
        'Insertar en la contabilidad
        sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) "
        sql = sql & " VALUES " & cad
        ConnConta.Execute sql
    Else
        cad = Format(Diario, "00") & ", " & DBSet(fecha, "F") & "," & Format(Asiento, "000000") & ","
        cad = cad & "0," & ValorNulo & "," & DBSet(Obs, "T")
        cad = "(" & cad & ")"
    
        'Insertar en la contabilidad
        sql = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) "
        sql = sql & " VALUES " & cad
        ConnConta.Execute sql
    End If
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabAsientoDia = False
        cadErr = Err.Description
    Else
        InsertarCabAsientoDia = True
    End If
End Function



Public Function InsertarLinAsientoDia(cad As String, cadErr As String) As Boolean
' el Tipo me indica desde donde viene la llamada
' tipo = 0 srecau.codmacta
' tipo = 1 scaalb.codmacta

Dim Rs As ADODB.Recordset
Dim Aux As String
Dim sql As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    If vParamAplic.ContabilidadNueva Then
        sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
        sql = sql & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
        sql = sql & " VALUES " & cad
    
    Else
 
        sql = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
        sql = sql & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
        sql = sql & " VALUES " & cad
        
    End If
    ConnConta.Execute sql

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinAsientoDia = False
        cadErr = Err.Description
    Else
        InsertarLinAsientoDia = True
    End If
End Function



Private Function EliminarCobros(cadErr As String) As Boolean

Dim Rs As ADODB.Recordset
Dim Aux As String
Dim sql As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    If vParamAplic.ContabilidadNueva Then
    
        sql = "DELETE FROM ariconta" & vParamAplic.NumeroConta & ".cobros where (numserie,numfactu,fecfactu,numorden) in "
        sql = sql & " (select nombre1, importe1, fecha1, importe2 from tmpinformes where codusu = " & vUsu.Codigo & ")"
    
    Else
 
        sql = "DELETE FROM conta" & vParamAplic.NumeroConta & ".scobro where (numserie,codfaccl,fecfaccl,numorden) in "
        sql = sql & " (select nombre1, importe1, fecha1, importe2 from tmpinformes where codusu = " & vUsu.Codigo & ")"
    
    End If
    
    conn.Execute sql

EInLinea:
    If Err.Number <> 0 Then
        EliminarCobros = False
        cadErr = Err.Description
    Else
        EliminarCobros = True
    End If
End Function



Private Function MarcarRegistros(cadErr As String) As Boolean

Dim Rs As ADODB.Recordset
Dim Aux As String
Dim sql As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

 
    sql = "UPDATE anecoop_pago SET idcontab = 1 where (expediente_id, expediente_pagoid) in "
    sql = sql & " (select nombre1, nombre2 from tmpinformes2 where codusu = " & vUsu.Codigo & ")"
    
    conn.Execute sql

EInLinea:
    If Err.Number <> 0 Then
        MarcarRegistros = False
        cadErr = Err.Description
    Else
        MarcarRegistros = True
    End If
End Function


