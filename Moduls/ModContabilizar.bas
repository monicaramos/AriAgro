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
Private RecImp As Currency
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




Public Function CrearTMPFacturas(cadTABLA As String, cadwhere As String) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturas = False
    
    Sql = "CREATE TEMPORARY TABLE tmpFactu ( "
    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
        Sql = Sql & "codtipom char(3) NOT NULL default '',"
        Sql = Sql & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        If cadTABLA = "scafpc" Or cadTABLA = "facturascom" Then
            Sql = Sql & "codprove int(6) unsigned NOT NULL default '0',"
            Sql = Sql & "numfactu varchar(10)  NOT NULL ,"
        Else
            Sql = Sql & "codtrans smallint(3) unsigned NOT NULL default '0',"
            Sql = Sql & "numfactu varchar(10)  NOT NULL ,"
        End If
    End If
    Sql = Sql & "fecfactu date NOT NULL default '0000-00-00') "
    conn.Execute Sql
     
     
    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
        Sql = "SELECT codtipom, numfactu, fecfactu"
    Else
        If cadTABLA = "scafpc" Or cadTABLA = "facturascom" Then
            Sql = "SELECT codprove, numfactu, fecfactu"
        Else
            Sql = "SELECT codtrans, numfactu, fecfactu"
        End If
    End If
    Sql = Sql & " FROM " & cadTABLA
    Sql = Sql & " WHERE " & cadwhere
    Sql = " INSERT INTO tmpFactu " & Sql
    conn.Execute Sql

    CrearTMPFacturas = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturas = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpFactu;"
        conn.Execute Sql
    End If
End Function


Public Sub BorrarTMPFacturas()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpFactu;"
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub InsertarTMPErrFac(MenError As String, cadwhere As String)
Dim Sql As String

    On Error Resume Next
    Sql = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
    Sql = Sql & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    Sql = Sql & " WHERE " & Replace(cadwhere, "scafpc", "tmpFactu")
    conn.Execute Sql
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Function CrearTMPErrFact(cadTABLA As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPErrFact = False
    
    Sql = "CREATE TEMPORARY TABLE tmpErrFac ( "
    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
        Sql = Sql & "codtipom char(3) NOT NULL default '',"
        Sql = Sql & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        Sql = Sql & "codprove int(6) unsigned NOT NULL default '0',"
        Sql = Sql & "numfactu varchar(10) NOT NULL ,"
    End If
    Sql = Sql & "fecfactu date NOT NULL default '0000-00-00', "
    Sql = Sql & "error varchar(200) NULL )"
    conn.Execute Sql
     
     CrearTMPErrFact = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrFact = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpErrFac;"
        conn.Execute Sql
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim Cad As String, devuelve As String

On Error GoTo EComprobarLetra

    ComprobarLetraSerie = False
    
    'Comprobar que existe la letra de serie en contabilidad
    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
        'cargamos el RSConta con la tabla contadores de BD: Contabilidad
        'donde estan todas las letra de serie que existen en la contabilidad
        Sql = "Select distinct tiporegi from contadores"
        Set RSconta = New ADODB.Recordset
        RSconta.Open Sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
        If RSconta.EOF Then
            RSconta.Close
            Set RSconta = Nothing
            Exit Function
        End If
            
    
        'obtenemos los distintos tipos de movimiento que vamos a contabilizar
        'de las facturas seleccionadas
        Sql = "select distinct " & cadTABLA & ".codtipom from " & cadTABLA
        Sql = Sql & " INNER JOIN tmpFactu ON " & cadTABLA & ".codtipom=tmpFactu.codtipom AND " & cadTABLA & ".numfactu=tmpFactu.numfactu AND " & cadTABLA & ".fecfactu=tmpFactu.fecfactu "
'        SQL = SQL & cadWHERE
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        b = True
        While Not Rs.EOF And b
            'comprobar que todas las letras serie existen en Ariges
'--monica:10/02/2009
'            SQL = "letraser"
'            devuelve = DevuelveDesdeBDNew(cAgro, "stipom", "codtipom", "codtipom", Rs!codTipoM, "T", SQL)
'++monica:10/02/2009
            Sql = ObtenerLetraSerie(Rs!codTipoM)
            devuelve = DBLet(Rs!codTipoM, "T")
'++
            If devuelve = "" Then
                b = False
                Cad = Rs!codTipoM & " en BD de Ariagro."
            ElseIf Sql <> "" Then
                'comprobar que todas las letras serie existen en la contabilidad
                devuelve = "tiporegi= " & DBSet(Sql, "T")
                RSconta.MoveFirst
                RSconta.Find (devuelve), , adSearchForward
                If RSconta.EOF Then
                    'no encontrado
                    b = False
                    Cad = Sql & " en BD de Contabilidad."
                End If
            End If
            If b Then Cad = Cad & DBSet(Rs!codTipoM, "T") & ","
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        RSconta.Close
        Set RSconta = Nothing
        
        If Not b Then 'Hay algun movimiento que no existe
            devuelve = "No existe el tipo de movimiento: " & Cad & vbCrLf
            devuelve = devuelve & "Consulte con el administrador."
            MsgBox devuelve, vbExclamation
            Exit Function
        End If
        
        'Todos los Tipo de movimiento existen
        If Cad <> "" Then
            Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitamos ult. coma
        
            'miramos si hay algun movimiento de factura que la letra serie sea nulo
            Sql = "select count(*) from usuarios.stipom "
            Sql = Sql & "where codtipom IN (" & Cad & ") and (isnull(letraser) or letraser='')"
            If RegistrosAListar(Sql) > 0 Then
                Sql = "Hay algun tipo de movimiento de Facturaci�n que no tiene letra serie." & vbCrLf
                Sql = Sql & "Comprobar en la tabla de tipos de movimiento: " & Cad
                MsgBox Sql, vbExclamation
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
Dim Sql As String
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
        Sql = "SELECT DISTINCT " & cadTABLA & ".codtipom,letraser," & cadTABLA & ".numfactu," & cadTABLA & ".fecfactu "
        Sql = Sql & " FROM (" & cadTABLA & " INNER JOIN usuarios.stipom stipom ON " & cadTABLA & ".codtipom=stipom.codtipom) "
        Sql = Sql & " INNER JOIN tmpFactu ON " & cadTABLA & ".codtipom=tmpFactu.codtipom AND " & cadTABLA & ".numfactu=tmpFactu.numfactu AND " & cadTABLA & ".fecfactu=tmpFactu.fecfactu "

        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF And b
            If vParamAplic.ContabilidadNueva Then
                Sql = "(numserie= " & DBSet(Rs!letraser, "T") & " AND numfactu=" & DBSet(Rs!NumFactu, "N") & " AND anofactu=" & Year(Rs!FecFactu) & ")"
            Else
                Sql = "(numserie= " & DBSet(Rs!letraser, "T") & " AND codfaccl=" & DBSet(Rs!NumFactu, "N") & " AND anofaccl=" & Year(Rs!FecFactu) & ")"
            End If
'            If SituarRSetMULTI(RSconta, SQL) Then
            Sql = SQLconta & Sql
            If RegistrosAListar(Sql, cConta) Then
                b = False
                Sql = "          Letra Serie: " & DBSet(Rs!letraser, "T") & vbCrLf
                Sql = Sql & "          N� Fac.: " & Format(Rs!NumFactu, "0000000") & vbCrLf
                Sql = Sql & "          Fecha: " & Rs!FecFactu
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            Sql = "Ya existe la factura: " & vbCrLf & Sql
            Sql = "Comprobando N� Facturas en Contabilidad...       " & vbCrLf & vbCrLf & Sql
            
            MsgBox Sql, vbExclamation
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
Dim Sql As String
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
        Sql = DevuelveDesdeBDNew(cConta, "parametros", "grupogto", "", "", "", cadG)
        If Sql <> "" And cadG <> "" Then
            Sql = " AND (codmacta like '" & Sql & "%' OR codmacta like '" & cadG & "%')"
        ElseIf Sql <> "" Then
            Sql = " AND (codmacta like '" & Sql & "%')"
        ElseIf cadG <> "" Then
            Sql = " AND (codmacta like '" & cadG & "%')"
        End If
        cadG = Sql
    End If
    
    
'    SQL = "SELECT codmacta FROM cuentas "
'    SQL = SQL & " WHERE apudirec='S'"
'    If cadG <> "" Then SQL = SQL & cadG
    
    SQLcuentas = "SELECT count(*) FROM cuentas WHERE apudirec='S' "
    If cadG <> "" Then SQLcuentas = SQLcuentas & cadG
    
    If Opcion = 1 Then
        If cadTABLA = "facturas" Then
            'Seleccionamos los distintos clientes,cuentas que vamos a facturar
            Sql = "SELECT DISTINCT facturas.codclien, clientes.codmacta "
            Sql = Sql & " FROM (facturas INNER JOIN clientes ON facturas.codclien=clientes.codclien) "
            Sql = Sql & " INNER JOIN tmpFactu ON facturas.codtipom=tmpFactu.codtipom AND facturas.numfactu=tmpFactu.numfactu AND facturas.fecfactu=tmpFactu.fecfactu "
        Else
            If cadTABLA = "facturassocio" Then
                SeccionHorto = DevuelveValor("select seccionhorto from rparam")
                'Seleccionamos las distintas cuentas de clientes de la seccion de horto, de los socios que vamos a facturar
                Sql = "SELECT DISTINCT facturassocio.codsocio, rsocios_seccion.codmaccli codmacta "
                Sql = Sql & " FROM (facturassocio INNER JOIN rsocios_seccion ON facturassocio.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & SeccionHorto & ") "
                Sql = Sql & " INNER JOIN tmpFactu ON facturassocio.codtipom=tmpFactu.codtipom AND facturassocio.numfactu=tmpFactu.numfactu AND facturassocio.fecfactu=tmpFactu.fecfactu "
            Else
                If cadTABLA = "scafpc" Then
                    'Seleccionamos los distintos proveedores,cuentas que vamos a facturar
                    Sql = "SELECT DISTINCT scafpc.codprove, proveedor.codmacta "
                    Sql = Sql & " FROM (scafpc INNER JOIN proveedor ON scafpc.codprove=proveedor.codprove) "
                    Sql = Sql & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
                Else
                    If cadTABLA = "facturascom" Then
                        'Seleccionamos los distintos proveedores,cuentas que vamos a facturar
                        Sql = "SELECT DISTINCT facturascom.codprove, proveedor.codmacta "
                        Sql = Sql & " FROM (facturascom INNER JOIN proveedor ON facturascom.codprove=proveedor.codprove) "
                        Sql = Sql & " INNER JOIN tmpFactu ON facturascom.codprove=tmpFactu.codprove AND facturascom.numfactu=tmpFactu.numfactu AND facturascom.fecfactu=tmpFactu.fecfactu "
                    Else
                        'Seleccionamos los distintos transportistas ,cuentas que vamos a facturar
                        Sql = "SELECT DISTINCT tcafpc.codtrans, agencias.codmacta "
                        Sql = Sql & " FROM (tcafpc INNER JOIN agencias ON tcafpc.codtrans=agencias.codtrans) "
                        Sql = Sql & " INNER JOIN tmpFactu ON tcafpc.codtrans=tmpFactu.codtrans AND tcafpc.numfactu=tmpFactu.numfactu AND tcafpc.fecfactu=tmpFactu.fecfactu "
                    End If
                End If
            End If
        End If
    ElseIf Opcion = 2 Or Opcion = 3 Or Opcion = 8 Then
        Sql = "SELECT distinct "
        If Opcion = 2 Then Sql = Sql & " sartic.codfamia,"
        If cadTABLA = "facturas" Then
            If Opcion <> 8 Then
                Sql = Sql & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((facturas_envases "
                Sql = Sql & " INNER JOIN tmpFactu ON facturas_envases.codtipom=tmpFactu.codtipom AND facturas_envases.numfactu=tmpFactu.numfactu AND facturas_envases.fecfactu=tmpFactu.fecfactu) "
                Sql = Sql & "INNER JOIN sartic ON facturas_envases.codartic=sartic.codartic) "
            Else
                numNivel = DevuelveDesdeBDNew(cConta, "empresa", "numnivel", "codempre", vParamAplic.NumeroConta, "N")
                NumDigit = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & numNivel, "codempre", vParamAplic.NumeroConta, "N")
                NumDigit3 = DevuelveDesdeBDNew(cConta, "empresa", "numdigi3", "codempre", vParamAplic.NumeroConta, "N")
                
'                CadCampo1 = "concat(concat(variedades.raizctavtas,tipomer.digicont), right(concat('0000000000',albaran_variedad.codvarie)," & (CCur(NumDigit) - CCur(NumDigit3) - 1) & "))"
                CadCampo1 = "CASE tipomer.tiptimer WHEN 0 THEN ctavtasinterior WHEN 1 THEN ctavtasexportacion WHEN 2 THEN ctavtasindustria WHEN 3 THEN ctavtasretirada WHEN 4 THEN ctavtasotros END"
                
                Sql = Sql & " albaran_variedad.codvarie, " & CadCampo1 & " as codmacta from ((((((facturas_variedad "
                Sql = Sql & " INNER JOIN tmpFactu ON facturas_variedad.codtipom=tmpFactu.codtipom AND facturas_variedad.numfactu=tmpFactu.numfactu AND facturas_variedad.fecfactu=tmpFactu.fecfactu) "
                Sql = Sql & " inner join usuarios.stipom stipom on facturas_variedad.codtipom=stipom.codtipom) "
                Sql = Sql & " inner join albaran on facturas_variedad.numalbar = albaran.numalbar) "
                Sql = Sql & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
                Sql = Sql & " inner join albaran_variedad on facturas_variedad.numalbar = albaran_variedad.numalbar and facturas_variedad.numlinealbar = albaran_variedad.numlinea) "
                Sql = Sql & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
                
                
'                Sql = Sql & " INNER JOIN tmpFactu ON facturas_variedad.codtipom=tmpFactu.codtipom AND facturas_variedad.numfactu=tmpFactu.numfactu AND facturas_variedad.fecfactu=tmpFactu.fecfactu) "
'                Sql = Sql & "INNER JOIN sartic ON facturas_envases.codartic=sartic.codartic) "
            End If
        Else
            If cadTABLA = "facturassocio" Then
                If Opcion <> 8 Then
                    Sql = Sql & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((facturassocio_envases "
                    Sql = Sql & " INNER JOIN tmpFactu ON facturassocio_envases.codtipom=tmpFactu.codtipom AND facturassocio_envases.numfactu=tmpFactu.numfactu AND facturassocio_envases.fecfactu=tmpFactu.fecfactu) "
                    Sql = Sql & "INNER JOIN sartic ON facturassocio_envases.codartic=sartic.codartic) "
                Else
                    numNivel = DevuelveDesdeBDNew(cConta, "empresa", "numnivel", "codempre", vParamAplic.NumeroConta, "N")
                    NumDigit = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & numNivel, "codempre", vParamAplic.NumeroConta, "N")
                    NumDigit3 = DevuelveDesdeBDNew(cConta, "empresa", "numdigi3", "codempre", vParamAplic.NumeroConta, "N")
                    
    '                CadCampo1 = "concat(concat(variedades.raizctavtas,tipomer.digicont), right(concat('0000000000',albaran_variedad.codvarie)," & (CCur(NumDigit) - CCur(NumDigit3) - 1) & "))"
                    CadCampo1 = "CASE tipomer.tiptimer WHEN 0 THEN ctavtasinterior WHEN 1 THEN ctavtasexportacion WHEN 2 THEN ctavtasindustria WHEN 3 THEN ctavtasretirada WHEN 4 THEN ctavtasotros END"
                    
                    Sql = Sql & " albaran_variedad.codvarie, " & CadCampo1 & " as codmacta from ((((((facturassocio_variedad "
                    Sql = Sql & " INNER JOIN tmpFactu ON facturassocio_variedad.codtipom=tmpFactu.codtipom AND facturassocio_variedad.numfactu=tmpFactu.numfactu AND facturassocio_variedad.fecfactu=tmpFactu.fecfactu) "
                    Sql = Sql & " inner join usuarios.stipom stipom on facturassocio_variedad.codtipom=stipom.codtipom) "
                    Sql = Sql & " inner join albaran on facturassocio_variedad.numalbar = albaran.numalbar) "
                    Sql = Sql & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
                    Sql = Sql & " inner join albaran_variedad on facturassocio_variedad.numalbar = albaran_variedad.numalbar and facturassocio_variedad.numlinealbar = albaran_variedad.numlinea) "
                    Sql = Sql & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
                End If
            Else
                If cadTABLA = "facturascom" Then
                    Sql = Sql & " variedades.ctacomtercero as codmacta,variedades.ctacomtercero as ctaabono from ((facturascom_variedad "
                    Sql = Sql & " INNER JOIN tmpFactu ON facturascom_variedad.codprove=tmpFactu.codprove AND facturascom_variedad.numfactu=tmpFactu.numfactu AND facturascom_variedad.fecfactu=tmpFactu.fecfactu) "
                    Sql = Sql & "INNER JOIN variedades ON facturascom_variedad.codvarie=variedades.codvarie) "
                Else
                    Sql = Sql & " sfamia.ctacompr as codmacta,sfamia.abocompr as ctaabono from ((slifpc "
                    Sql = Sql & " INNER JOIN tmpFactu ON slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu) "
                    Sql = Sql & "INNER JOIN sartic ON slifpc.codartic=sartic.codartic) "
                End If
            End If
        End If
        If Opcion <> 8 And cadTABLA <> "facturascom" Then Sql = Sql & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
    ElseIf Opcion = 4 Or Opcion = 6 Then
        Sql = "select distinct " & DBSet(vParamAplic.CtaTraReten, "T") & " as codmacta from tcafpc "
    ElseIf Opcion = 5 Or Opcion = 7 Then
'        Sql = "select distinct " & DBSet(vParamAplic.CtaAboTrans, "T") & " as codmacta from tcafpc "
'       transporte
            Sql = " SELECT if(tipomer.tiptimer = 1,variedades.ctatraexporta,variedades.ctatrainterior) as cuenta "
            Sql = Sql & " FROM tlifpc, albaran, albaran_variedad, variedades, tipomer, tmpFactu, tcafpc  WHERE "
            Sql = Sql & " tcafpc.tipo = 0 and " ' transportista
            Sql = Sql & " tlifpc.codtrans=tmpFactu.codtrans and tlifpc.numfactu=tmpFactu.numfactu and tlifpc.fecfactu=tmpFactu.fecfactu and "
            Sql = Sql & " tlifpc.numalbar=albaran_variedad.numalbar and "
            Sql = Sql & " tlifpc.numlinea=albaran_variedad.numlinea and "
            Sql = Sql & " tlifpc.codtrans=tcafpc.codtrans and tlifpc.numfactu=tcafpc.numfactu and tlifpc.fecfactu=tcafpc.fecfactu and "
            Sql = Sql & " albaran_variedad.numalbar=albaran.numalbar and "
            Sql = Sql & " albaran_variedad.codvarie=variedades.codvarie and "
            Sql = Sql & " albaran.codtimer=tipomer.codtimer "
            Sql = Sql & " group by 1 "

    ElseIf Opcion = 12 Or Opcion = 13 Then
'       comisionista
            Sql = " SELECT variedades.ctacomisionista as cuenta, variedades.codvarie  "
            Sql = Sql & " FROM tlifpc, albaran, albaran_variedad, variedades, tipomer, tmpFactu, tcafpc  WHERE "
            Sql = Sql & " tcafpc.tipo = 1 and " ' comisionista
            Sql = Sql & " tlifpc.codtrans=tmpFactu.codtrans and tlifpc.numfactu=tmpFactu.numfactu and tlifpc.fecfactu=tmpFactu.fecfactu and "
            Sql = Sql & " tlifpc.numalbar=albaran_variedad.numalbar and "
            Sql = Sql & " tlifpc.numlinea=albaran_variedad.numlinea and "
            Sql = Sql & " tlifpc.codtrans=tcafpc.codtrans and tlifpc.numfactu=tcafpc.numfactu and tlifpc.fecfactu=tcafpc.fecfactu and "
            Sql = Sql & " albaran_variedad.numalbar=albaran.numalbar and "
            Sql = Sql & " albaran_variedad.codvarie=variedades.codvarie and "
            Sql = Sql & " albaran.codtimer=tipomer.codtimer "
            Sql = Sql & " group by 1 "
            
    ElseIf Opcion = 9 Or Opcion = 10 Then
            Sql = " select codmacta as cuenta "
            Sql = Sql & " from tcafpv, tmpFactu "
            Sql = Sql & " where tmpFactu.codtrans=tcafpv.codtrans and tmpFactu.numfactu=tcafpv.numfactu and tmpFactu.fecfactu=tcafpv.fecfactu "
            Sql = Sql & " group by 1 "
    ElseIf Opcion = 14 Then
            '[Monica]07/03/2019: facturas de compra de variedades
            Sql = " SELECT variedades.ctacomtercero as cuenta, variedades.codvarie  "
            Sql = Sql & " FROM facturascom_variedad, variedades, tmpFactu WHERE "
            Sql = Sql & " facturascom_variedad.codprove=tmpFactu.codprove and facturascom_variedad.numfactu=tmpFactu.numfactu and facturascom_variedad.fecfactu=tmpFactu.fecfactu and "
            Sql = Sql & " facturascom_variedad.codvarie=variedades.codvarie "
            Sql = Sql & " group by 1 "
    
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    b = True

    While Not Rs.EOF And b
        If Opcion < 4 Or Opcion = 8 Then
            Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!Codmacta, "T")
        ElseIf Opcion = 4 Or Opcion = 6 Then
            Sql = SQLcuentas & " AND codmacta= " & DBSet(vParamAplic.CtaTraReten, "T")
        ElseIf Opcion = 5 Or Opcion = 7 Then
            Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!Cuenta, "T")
        ElseIf Opcion = 12 Or Opcion = 13 Or Opcion = 14 Then
            Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!Cuenta, "T")
        ElseIf Opcion = 9 Or Opcion = 10 Then
            Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!Cuenta, "T")
        End If
            
        
        If Not (RegistrosAListar(Sql, cConta) > 0) Then
        'si no lo encuentra
            b = False 'no encontrado
            If Opcion = 1 Then
                If cadTABLA = "facturas" Then
                    Sql = Rs!Codmacta & " del Cliente " & Format(Rs!CodClien, "000000")
                Else
                    If cadTABLA = "facturassocio" Then
                        Sql = Rs!Codmacta & " del Socio " & Format(Rs!CodSocio, "000000")
                    Else
                        If cadTABLA = "scafpc" Or cadTABLA = "facturascom" Then
                            Sql = Rs!Codmacta & " del Proveedor " & Format(Rs!codProve, "000000")
                        Else
                            Sql = Rs!Codmacta & " del Transportista " & Format(Rs!codTrans, "000")
                        End If
                    End If
                End If
            ElseIf Opcion = 2 Then
                Sql = Rs!Codmacta & " de la familia " & Format(Rs!codfamia, "0000")
            ElseIf Opcion = 3 Then
                Sql = Rs!Codmacta
            ElseIf Opcion = 4 Or Opcion = 6 Then
                Sql = vParamAplic.CtaTraReten
            ElseIf Opcion = 5 Or Opcion = 7 Then
                Sql = DBLet(Rs!Cuenta, "T") ' vParamAplic.CtaAboTrans
            ElseIf Opcion = 12 Or Opcion = 13 Then
                Sql = DBLet(Rs!Cuenta, "T") & " de comisionista de la variedad " & Format(Rs!codvarie, "000000")
            ElseIf Opcion = 14 Then
                Sql = DBLet(Rs!Cuenta, "T") & " de compras terceros de la variedad " & Format(Rs!codvarie, "000000")
            ElseIf Opcion = 8 Then
                Sql = Rs!Codmacta & " de la variedad " & Format(Rs!codvarie, "0000")
            ElseIf Opcion = 9 Or Opcion = 10 Then
                Sql = DBLet(Rs!Cuenta, "T") ' vParamAplic.CtaAboTrans
            End If
        End If
        
        
        If Opcion = 2 Or Opcion = 3 Then
            'Comprobar que ademas de existir la cuenta de ventas exista tambien
            'la cuenta ABONO ventas (sfamia.aboventa)
            '---------------------------------------------
            Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!ctaabono, "T")
'            RSconta.MoveFirst
'            RSconta.Find (SQL), , adSearchForward
'            If RSconta.EOF Then
            If Not (RegistrosAListar(Sql, cConta) > 0) Then
                b = False 'no encontrado
                If Opcion = 2 Then
                    Sql = Rs!ctaabono & " de la familia " & Format(Rs!codfamia, "0000")
                ElseIf Opcion = 3 Then
                    Sql = Rs!ctaabono
                End If
            End If
            
            
            'comprobar cuentas alternativas solo para facturacion a CLIENTES
            '----------------------------------------------------------------
            If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
                ' Comprobar cuenta VENTA alternativa
                If DBLet(Rs!ctavent1, "T") <> "" Then
                    Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!ctavent1, "T")
'                    RSconta.MoveFirst
'                    RSconta.Find (SQL), , adSearchForward
'                    If RSconta.EOF Then
                    If Not (RegistrosAListar(Sql, cConta) > 0) Then
                        b = False 'no encontrado
                        If Opcion = 2 Then
                            Sql = Rs!ctavent1 & " de la familia " & Format(Rs!codfamia, "0000")
                        ElseIf Opcion = 3 Then
                            Sql = Rs!ctavent1
                        End If
                    End If
                Else
                    b = False
                    Sql = " o la familia no tiene asignada cuenta venta alternativa."
                End If
                
                ' Comprobar cuenta de ABONO alternativa
                If DBLet(Rs!abovent1, "T") <> "" Then
                    Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!abovent1, "T")
'                    RSconta.MoveFirst
'                    RSconta.Find (SQL), , adSearchForward
'                    If RSconta.EOF Then
                    If Not (RegistrosAListar(Sql, cConta) > 0) Then
                        b = False 'no encontrado
                        If Opcion = 2 Then
                            Sql = Rs!abovent1 & " de la familia " & Format(Rs!codfamia, "0000")
                        ElseIf Opcion = 3 Then
                            Sql = Rs!abovent1
                        End If
                    End If
                Else
                    b = False
                    Sql = " o la familia no tiene asignada cuenta abono alternativa."
                End If
            End If
            
        End If
        
        Rs.MoveNext
    Wend
    
    
        
        
        If Not b Then
            If Not (Opcion = 3 Or Opcion = 6 Or Opcion = 7) Then
                Sql = "No existe la cta contable " & Sql
            Else
                Sql = "La cuenta " & Sql & " no es del nivel correcto. "
                If Opcion = 3 Then Sql = Sql & "(Familias de art�culos)."
            End If
            Sql = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & Sql
            
            MsgBox Sql, vbExclamation
            ComprobarCtaContable_new = False
        Else
            ComprobarCtaContable_new = True
        End If
    Exit Function
    
ECompCta:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function







Public Function ComprobarTiposIVA(cadTABLA As String) As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (scafac.codigiva1, codigiv2,codigiv3)
'que vamos a contabilizar existan en la contabilidad
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVA = False
    
    Sql = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        For i = 1 To 3
            If cadTABLA = "facturas" Then
                Sql = "SELECT DISTINCT facturas.codiiva" & i
                Sql = Sql & " FROM facturas "
                Sql = Sql & " INNER JOIN tmpFactu ON facturas.codtipom=tmpFactu.codtipom AND facturas.numfactu=tmpFactu.numfactu AND facturas.fecfactu=tmpFactu.fecfactu "
                Sql = Sql & " WHERE not isnull(codiiva" & i & ")"
            Else
                If cadTABLA = "facturassocio" Then
                    Sql = "SELECT DISTINCT facturassocio.codiiva" & i
                    Sql = Sql & " FROM facturassocio "
                    Sql = Sql & " INNER JOIN tmpFactu ON facturassocio.codtipom=tmpFactu.codtipom AND facturassocio.numfactu=tmpFactu.numfactu AND facturassocio.fecfactu=tmpFactu.fecfactu "
                    Sql = Sql & " WHERE not isnull(codiiva" & i & ")"
                Else
                    If cadTABLA = "scafpc" Then
                        Sql = "SELECT DISTINCT scafpc.tipoiva" & i
                        Sql = Sql & " FROM " & cadTABLA
                        Sql = Sql & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
                        Sql = Sql & " WHERE not isnull(tipoiva" & i & ")"
                    Else
                        If cadTABLA = "facturascom" Then
                            Sql = "SELECT DISTINCT facturascom.codiiva" & i
                            Sql = Sql & " FROM " & cadTABLA
                            Sql = Sql & " INNER JOIN tmpFactu ON facturascom.codprove=tmpFactu.codprove AND facturascom.numfactu=tmpFactu.numfactu AND facturascom.fecfactu=tmpFactu.fecfactu "
                            Sql = Sql & " WHERE not isnull(codiiva" & i & ")"
                        Else
                            Sql = "SELECT DISTINCT tcafpc.tipoiva" & i
                            Sql = Sql & " FROM " & cadTABLA
                            Sql = Sql & " INNER JOIN tmpFactu ON tcafpc.codtrans=tmpFactu.codtrans AND tcafpc.numfactu=tmpFactu.numfactu AND tcafpc.fecfactu=tmpFactu.fecfactu "
                            Sql = Sql & " WHERE not isnull(tipoiva" & i & ")"
                        End If
                    End If
                End If
            End If
'            SQL = SQL & " WHERE " & cadWHERE & " AND codigiv" & i & " <> 0 "

            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            b = True
            While Not Rs.EOF And b
                Sql = "codigiva= " & DBSet(Rs.Fields(0), "N")
                RSconta.MoveFirst
                RSconta.Find (Sql), , adSearchForward
                If RSconta.EOF Then
                    b = False 'no encontrado
                    Sql = "Tipo de IVA: " & Rs.Fields(0)
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        
            If Not b Then
                Sql = "No existe el " & Sql
                Sql = "Comprobando Tipos de IVA en contabilidad..." & vbCrLf & vbCrLf & Sql
            
                MsgBox Sql, vbExclamation
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
Dim Sql As String

    On Error GoTo ECCoste

    ComprobarCCoste = False
    Sql = vUsu.Login
    If Sql <> "" Then
        cadCC = DevuelveDesdeBDNew(cAgro, "straba", "codccost", "login", Sql, "T")
        If cadCC <> "" Then
            'comprobar que el Centro de Coste existe en la Contabilidad
            If vParamAplic.ContabilidadNueva Then
                Sql = DevuelveDesdeBDNew(cConta, "ccoste", "codccost", "codccost", cadCC, "T")
            Else
                Sql = DevuelveDesdeBDNew(cConta, "cabccost", "codccost", "codccost", cadCC, "T")
            End If
            If Sql <> "" Then
                ComprobarCCoste = True
            Else
                Sql = "No existe el CC: " & cadCC
                Sql = "Comprobando Centros de Coste en contabilidad..." & vbCrLf & vbCrLf & Sql
                MsgBox Sql, vbExclamation
            End If
        Else 'el usuario no tiene asignado un centro de coste
            Sql = "El trabajador conectado no tiene asignado un centro de coste."
            Sql = "Comprobando Centros de Coste ..." & vbCrLf & vbCrLf & Sql
            MsgBox Sql, vbExclamation
        End If
    End If
    
ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Centros de Coste", Err.Description
    End If
End Function


Public Function ComprobarCCoste_new(cadCC As String, cadTABLA As String, Optional Opcion As Byte) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECCoste

    ComprobarCCoste_new = False
    Select Case cadTABLA
        Case "facturas" ' facturas de venta
            Select Case Opcion
                Case 1
                    Sql = "select distinct variedades.codccost from facturas_variedad, albaran_variedad, variedades, tmpFactu where "
                    Sql = Sql & " albaran_variedad.codvarie=variedades.codvarie and "
                    Sql = Sql & " facturas_variedad.codtipom=tmpFactu.codtipom AND facturas_variedad.numfactu=tmpFactu.numfactu AND facturas_variedad.fecfactu=tmpFactu.fecfactu and  "
                    Sql = Sql & " albaran_variedad.numalbar = facturas_variedad.numalbar and "
                    Sql = Sql & " albaran_variedad.numlinea = facturas_variedad.numlinealbar "
                Case 2
                    Sql = " select distinct sfamia.codccost from facturas_envases, sartic, sfamia, tmpFactu where "
                    Sql = Sql & " facturas_envases.codtipom=tmpFactu.codtipom AND facturas_envases.numfactu=tmpFactu.numfactu AND facturas_envases.fecfactu=tmpFactu.fecfactu and  "
                    Sql = Sql & " facturas_envases.codartic = sartic.codartic and "
                    Sql = Sql & " sartic.codfamia = sfamia.codfamia "
                Case 3
                    If HayFacturasACuenta Then
                        Sql = " select '" & vParamAplic.CCosteFraACta & "' as codccost from tmpFactu where tmpfactu.codtipom = 'EAC' "
                    Else
                        ComprobarCCoste_new = True
                        Exit Function
                    End If
            End Select
        
        Case "facturassocio" ' facturas de venta a socios
            Select Case Opcion
                Case 1
                    Sql = "select distinct variedades.codccost from facturassocio_variedad, albaran_variedad, variedades, tmpFactu where "
                    Sql = Sql & " albaran_variedad.codvarie=variedades.codvarie and "
                    Sql = Sql & " facturassocio_variedad.codtipom=tmpFactu.codtipom AND facturassocio_variedad.numfactu=tmpFactu.numfactu AND facturassocio_variedad.fecfactu=tmpFactu.fecfactu and  "
                    Sql = Sql & " albaran_variedad.numalbar = facturassocio_variedad.numalbar and "
                    Sql = Sql & " albaran_variedad.numlinea = facturassocio_variedad.numlinealbar "
                Case 2
                    Sql = " select distinct sfamia.codccost from facturassocio_envases, sartic, sfamia, tmpFactu where "
                    Sql = Sql & " facturassocio_envases.codtipom=tmpFactu.codtipom AND facturassocio_envases.numfactu=tmpFactu.numfactu AND facturassocio_envases.fecfactu=tmpFactu.fecfactu and  "
                    Sql = Sql & " facturassocio_envases.codartic = sartic.codartic and "
                    Sql = Sql & " sartic.codfamia = sfamia.codfamia "
                Case 3
                    If HayFacturasACuenta Then
                        Sql = " select '" & vParamAplic.CCosteFraACta & "' as codccost from tmpFactu where tmpfactu.codtipom = 'EAC' "
                    Else
                        ComprobarCCoste_new = True
                        Exit Function
                    End If
            End Select
        
        Case "scafpc" ' facturas de compra
            Sql = " select distinct sfamia.codccost from slifpc, sartic, sfamia, tmpFactu where "
            Sql = Sql & " slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu and  "
            Sql = Sql & " slifpc.codartic = sartic.codartic and "
            Sql = Sql & " sartic.codfamia = sfamia.codfamia "
        
        Case "tcafpc" ' facturas de transporte
            Sql = "select distinct variedades.codccost from tlifpc, albaran_variedad, variedades, tmpFactu where "
            Sql = Sql & " albaran_variedad.codvarie=variedades.codvarie and "
            Sql = Sql & " tlifpc.codtrans=tmpFactu.codtrans AND tlifpc.numfactu=tmpFactu.numfactu AND tlifpc.fecfactu=tmpFactu.fecfactu and  "
            Sql = Sql & " albaran_variedad.numalbar = tlifpc.numalbar and "
            Sql = Sql & " albaran_variedad.numlinea = tlifpc.numlinea "
    
        Case "facturascom" ' facturas de compra
            Sql = "select distinct variedades.codccost from facturascom_variedad, variedades, tmpFactu where "
            Sql = Sql & " facturascom_variedad.codvarie=variedades.codvarie and "
            Sql = Sql & " facturascom_variedad.codprove=tmpFactu.codprove AND facturascom_variedad.numfactu=tmpFactu.numfactu AND facturascom_variedad.fecfactu=tmpFactu.fecfactu  "
    
    
    
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = True

    While Not Rs.EOF And b
        If DBLet(Rs.Fields(0).Value, "T") = "" Then
            b = False
        Else
            If vParamAplic.ContabilidadNueva Then
                Sql = DevuelveDesdeBDNew(cConta, "ccoste", "codccost", "codccost", Rs.Fields(0).Value, "T")
            Else
                Sql = DevuelveDesdeBDNew(cConta, "cabccost", "codccost", "codccost", Rs.Fields(0).Value, "T")
            End If
            If Sql = "" Then
                b = False
                Sql2 = "Centro de Coste: " & Rs.Fields(0)
            End If
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
    If Not b Then
        Sql = "No existe el " & Sql2
        Sql = "Comprobando Centros de Coste en contabilidad..." & vbCrLf & vbCrLf & Sql
    
        MsgBox Sql, vbExclamation
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
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECCoste

    ComprobarFormadePago = False
    Sql = "select distinct " & cTabla & ".codforpa from " & cTabla & ", tmpFactu where "
    Sql = Sql & cTabla & ".codtipom=tmpFactu.codtipom AND " & cTabla & ".numfactu=tmpFactu.numfactu AND " & cTabla & ".fecfactu=tmpFactu.fecfactu  "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = True

    While Not Rs.EOF And b
        If Not vParamAplic.ContabilidadNueva Then
            Sql = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", Rs.Fields(0).Value, "N")
        Else
            Sql = DevuelveDesdeBDNew(cConta, "formapago", "codforpa", "codforpa", Rs.Fields(0).Value, "N")
        End If
        If Sql = "" Then
            b = False
            Sql2 = "Formas de Pago: " & Rs.Fields(0)
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
    If Not b Then
        Sql = "No existe la " & Sql2
        Sql = "Comprobando Formas de Pago en contabilidad..." & vbCrLf & vbCrLf & Sql
    
        MsgBox Sql, vbExclamation
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




Public Function PasarFactura(cadwhere As String, CodCCost As String, CtaBan As String, cTabla As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariges.scafac --> conta.cabfact
' ariges.slifac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim Sql As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    '$$$
    'Insertar en la conta Cabecera Factura
    b = InsertarCabFact(cTabla, cadwhere, cadMen, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        CCoste = CodCCost
        'Insertar lineas de Factura en la Conta
        If vParamAplic.ContabilidadNueva Then
            b = InsertarLinFact_newContaNueva(cTabla, cadwhere, cadMen)
        Else
            b = InsertarLinFact_new(cTabla, cadwhere, cadMen)
        End If
        cadMen = "Insertando Lin. Factura: " & cadMen

        '++monica:a�adida la parte de insertar en tesoreria
        If b Then
            Select Case cTabla
                Case "facturas"
                    b = InsertarEnTesoreriaNewFac(cadwhere, CtaBan, cadMen)
                Case "facturassocio"
                    b = InsertarEnTesoreriaNewFacSoc(cadwhere, CtaBan, cadMen)
            End Select
            cadMen = "Insertando en Tesoreria: " & cadMen
        End If
        
        '++


        If b Then
            If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
        
            'Poner intconta=1 en ariagro.facturas
            b = ActualizarCabFact(cTabla, cadwhere, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
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
        
        Sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        Sql = Sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
        Sql = Sql & " WHERE " & Replace(cadwhere, cTabla, "tmpFactu")
        conn.Execute Sql
    End If
End Function


Private Function InsertarCabFact(cTabla As String, cadwhere As String, cadErr As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Intracom As Integer
Dim SeccionHorto As Integer
Dim CadenaInsertFaclin2 As String


    On Error GoTo EInsertar
    Sql = ""
    Select Case cTabla
        Case "facturas"
            Sql = Sql & " SELECT stipom.letraser,numfactu,fecfactu, clientes.codmacta,clientes.cliabono,year(fecfactu) as anofaccl,"
            Sql = Sql & "baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
            Sql = Sql & "totalfac,codiiva1,codiiva2,codiiva3, porcrec1, porcrec2, porcrec3, imporec1, imporec2, imporec3, clientes.codpaise "
            Sql = Sql & ",nomclien,domclien,codpobla,pobclien,proclien,cifclien,facturas.codforpa"
            Sql = Sql & " FROM (" & "facturas inner join " & "usuarios.stipom stipom on facturas.codtipom=stipom.codtipom) "
            Sql = Sql & "INNER JOIN " & "clientes ON facturas.codclien=clientes.codclien "
            Sql = Sql & " WHERE " & cadwhere
    
        Case "facturassocio"
            SeccionHorto = DevuelveValor("select seccionhorto from rparam")
        
            Sql = Sql & " SELECT stipom.letraser,numfactu,fecfactu, rsocios_seccion.codmaccli codmacta,false cliabono,year(fecfactu) as anofaccl,"
            Sql = Sql & "baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
            Sql = Sql & "totalfac,codiiva1,codiiva2,codiiva3, porcrec1, porcrec2, porcrec3, imporec1, imporec2, imporec3, 0 codpaise "
            Sql = Sql & ",nomsocio nomclien,dirsocio domclien,codpostal codpobla,pobsocio pobclien,prosocio proclien,nifsocio cifclien,facturassocio.codforpa "
            Sql = Sql & " FROM ((" & "facturassocio inner join " & "usuarios.stipom stipom on facturassocio.codtipom=stipom.codtipom) "
            Sql = Sql & " INNER JOIN rsocios on facturassocio.codsocio = rsocios.codsocio) "
            Sql = Sql & "INNER JOIN " & "rsocios_seccion ON facturassocio.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(SeccionHorto, "N")
            Sql = Sql & " WHERE " & cadwhere
    
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
    
        vContaFra.NumeroFactura = DBLet(Rs!NumFactu)
        vContaFra.Anofac = Year(DBLet(Rs!FecFactu))
        vContaFra.Serie = DBLet(Rs!letraser)
        
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = 0
        DtoGnral = 0
        BaseImp = Rs!baseimp1 + CCur(DBLet(Rs!baseimp2, "N")) + CCur(DBLet(Rs!baseimp3, "N"))
        IvaImp = DBLet(Rs!impoiva1, "N") + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
        RecImp = DBLet(Rs!imporec1, "N") + DBLet(Rs!imporec2, "N") + DBLet(Rs!imporec3, "N")
        '---- Laura 10/10/2006:  a�adir el totalfac para utilizarlo en insertar lineas
        TotalFac = Rs!TotalFac
        '----
        conCtaAlt = Rs!cliAbono

        Intracom = 0
        If Not DBSet(Rs!CodPaise, "N", "S") = ValorNulo Then
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "paises", "intracom", "codpaise", Rs!CodPaise, "N")
            If Sql <> "" Then Intracom = CInt(Sql)
        End If
        
        If vParamAplic.ContabilidadNueva Then
            Sql = ""
            Sql = DBSet(Rs!letraser, "T") & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!FecFactu) & ",'FACTURACION',"
            
            ' para el caso de las rectificativas
            Dim vTipM As String
            vTipM = DevuelveValor("select codtipom from usuarios.stipom where letraser = " & DBSet(Rs!letraser, "T"))
            
            '[Monica]30/05/2017: para el caso de q sea una intracomunitaria de venta el tipo de intracomunitaria es E
            If Intracom = 1 Then
            
                '[Monica]27/07/2017: antes estaba esto
                'SQL = SQL & "'0',1,'E',"
                If Not IsNull(Rs!porciva2) Then
                    Sql = Sql & "'C',"
                Else
                    Sql = Sql & "'0',"
                End If
                Sql = Sql & "1,'E',"
            Else
                If vTipM = "FAR" Then
                    Sql = Sql & "'D',"
                Else
                    '[Monica]27/07/2017: antes si o si era 0
                    If Not IsNull(Rs!porciva2) Then
                        Sql = Sql & "'C',"
                    Else
                        Sql = Sql & "'0',"
                    End If
                End If
                '[Monica]12/11/2018: quedan los casos: 0=general o 2=exportacion-importacion
                'Sql = Sql & "0," & ValorNulo & ","
                Sql = Sql & DBSet(Intracom, "N") & "," & ValorNulo & ","
            End If
            
            Dim LetraPais As String
            
            LetraPais = DevuelveDesdeBDNew(cAgro, "paises", "letraspais", "codpaise", DBLet(Rs!CodPaise, "N"), "N")
            If LetraPais = "" Then LetraPais = "ES"
            
            Sql = Sql & DBSet(Rs!Codforpa, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
            Sql = Sql & DBSet(RecImp, "N") & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!FecFactu, "F") & ","
            Sql = Sql & DBSet(Rs!Nomclien, "T") & "," & DBSet(Rs!domclien, "T") & "," & DBSet(Rs!codPobla, "T") & "," & DBSet(Rs!pobclien, "T") & ","
            Sql = Sql & DBSet(Rs!proclien, "T") & "," & DBSet(Rs!cifClien, "T") & "," & DBSet(LetraPais, "T") & ",1"
            
            Cad = "(" & Sql & ")"
        Else
            Sql = ""
            Sql = "'" & Rs!letraser & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!FecFactu) & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!baseimp1, "N") & "," & DBSet(Rs!baseimp2, "N", "S") & "," & DBSet(Rs!baseimp3, "N", "S") & "," & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", "S") & "," & DBSet(Rs!porciva3, "N", "S") & ","
            Sql = Sql & DBSet(Rs!porcrec1, "N", "S") & "," & DBSet(Rs!porcrec2, "N", "S") & "," & DBSet(Rs!porcrec3, "N", "S") & "," & DBSet(Rs!impoiva1, "N", "N") & "," & DBSet(Rs!impoiva2, "N", "S") & "," & DBSet(Rs!impoiva3, "N", "S") & ","
            Sql = Sql & DBSet(Rs!imporec1, "N", "S") & "," & DBSet(Rs!imporec2, "N", "S") & "," & DBSet(Rs!imporec3, "N", "S") & ","
            Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!codiiva1, "N") & "," & DBSet(Rs!codiiva2, "N", "S") & "," & DBSet(Rs!codiiva3, "N", "S") & "," & DBSet(Intracom, "N") & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!FecFactu, "F")
            Cad = "(" & Sql & ")"
        End If
'        RS.MoveNext

    End If
    
    
    If vParamAplic.ContabilidadNueva Then
        Sql = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,codconce340,codopera,codintra,codforpa,totbases,totbasesret,totivas,"
        Sql = Sql & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
        Sql = Sql & "codpais,codagente)"
        Sql = Sql & " VALUES " & Cad
        ConnConta.Execute Sql
'***
        CadenaInsertFaclin2 = ""
            
        
        'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
        'IVA 1, siempre existe
        Sql2 = "'" & Rs!letraser & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
        Sql2 = Sql2 & "1," & DBSet(Rs!baseimp1, "N") & "," & Rs!codiiva1 & "," & DBSet(Rs!porciva1, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!porcrec1, "N", "S") & "," & DBSet(Rs!impoiva1, "N") & "," & DBSet(Rs!imporec1, "N", "S")
        CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
        
        'para las lineas
        vTipoIva(0) = Rs!codiiva1
        vPorcIva(0) = Rs!porciva1
        '[Monica]20/12/2018: se mandaba un 0
        vPorcRec(0) = DBLet(Rs!porcrec1, "N") '0
        vImpIva(0) = Rs!impoiva1
        '[Monica]20/12/2018: se mandaba un 0
        vImpRec(0) = DBLet(Rs!imporec1, "N") '0
        
        vBaseIva(0) = Rs!baseimp1
        
        vTipoIva(1) = 0: vTipoIva(2) = 0
        
        If Not IsNull(Rs!porciva2) Then
            Sql2 = "'" & Rs!letraser & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
            Sql2 = Sql2 & "2," & DBSet(Rs!baseimp2, "N") & "," & Rs!codiiva2 & "," & DBSet(Rs!porciva2, "N") & ","
            Sql2 = Sql2 & DBSet(Rs!porcrec2, "N", "S") & "," & DBSet(Rs!impoiva2, "N") & "," & DBSet(Rs!imporec2, "N", "S")
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
            vTipoIva(1) = Rs!codiiva2
            vPorcIva(1) = Rs!porciva2
            vPorcRec(1) = DBLet(Rs!porcrec2, "N") '0
            vImpIva(1) = Rs!impoiva2
            vImpRec(1) = DBLet(Rs!imporec2, "N") '0
            vBaseIva(1) = Rs!baseimp2
        End If
        If Not IsNull(Rs!porciva3) Then
            Sql2 = "'" & Rs!letraser & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
            Sql2 = Sql2 & "3," & DBSet(Rs!baseimp3, "N") & "," & Rs!codiiva3 & "," & DBSet(Rs!porciva3, "N") & ","
            Sql2 = Sql2 & DBSet(Rs!porcrec3, "N", "S") & "," & DBSet(Rs!impoiva3, "N") & "," & DBSet(Rs!imporec3, "N", "S")
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
            vTipoIva(2) = Rs!codiiva3
            vPorcIva(2) = Rs!porciva3
            vPorcRec(2) = DBLet(Rs!porcrec3, "N") '0
            vImpIva(2) = Rs!impoiva3
            vImpRec(2) = DBLet(Rs!imporec3, "N") '0
            vBaseIva(2) = Rs!baseimp3
        End If

        Sql = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
        Sql = Sql & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
        ConnConta.Execute Sql
'***
    Else
        'Insertar en la contabilidad
        Sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
        Sql = Sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
        Sql = Sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
        Sql = Sql & " VALUES " & Cad
    
        ConnConta.Execute Sql
    
    
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



Private Function InsertarLinFact(cadTABLA As String, cadwhere As String, cadErr As String, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    If cadTABLA = "scafac" Then
        Sql = " SELECT stipom.letraser,slifac.codtipom,numfactu,fecfactu,sartic.codfamia,sfamia.ctaventa,sfamia.ctavent1,sfamia.aboventa,sfamia.abovent1,sum(importel) as importe "
        Sql = Sql & " FROM ((slifac inner join usuarios.stipom stipom on slifac.codtipom=stipom.codtipom) "
        Sql = Sql & " inner join sartic on slifac.codartic=sartic.codartic) "
        Sql = Sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        Sql = Sql & " WHERE " & Replace(cadwhere, "scafac", "slifac")
        Sql = Sql & " GROUP BY sfamia.codfamia "
    Else
        Sql = " SELECT slifpc.codprove,numfactu,fecfactu,sartic.codfamia,sfamia.ctacompr,sfamia.abocompr,sum(importel) as importe "
        Sql = Sql & " FROM (slifpc  "
        Sql = Sql & " inner join sartic on slifpc.codartic=sartic.codartic) "
        Sql = Sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        Sql = Sql & " WHERE " & Replace(cadwhere, "scafpc", "slifpc")
        Sql = Sql & " GROUP BY sfamia.codfamia "
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = Cad
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
        Sql = ""
        Sql2 = ""
        If cadTABLA = "scafac" Then
            Sql = "'" & Rs!letraser & "'," & Rs!NumFactu & "," & Year(Rs!FecFactu) & "," & i & ","
            If Not conCtaAlt Then 'cliente no tiene cuenta alternativa
                If ImpLinea >= 0 Then
                    Sql = Sql & DBSet(Rs!ctaventa, "T")
                Else
                    Sql = Sql & DBSet(Rs!aboventa, "T")
                End If
            Else
                If ImpLinea >= 0 Then
                    Sql = Sql & DBSet(Rs!ctavent1, "T")
                Else
                    Sql = Sql & DBSet(Rs!abovent1, "T")
                End If
            End If
        Else
            Sql = numRegis & "," & Year(Rs!FecFactu) & "," & i & ","
            If ImpLinea >= 0 Then
                Sql = Sql & DBSet(Rs!ctacompr, "T")
            Else
                Sql = Sql & DBSet(Rs!abocompr, "T")
            End If
        End If
        Sql2 = Sql & ","
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        If CCoste = "" Then
            Sql = Sql & ValorNulo
        Else
            Sql = Sql & DBSet(CCoste, "T")
        End If
        
        Cad = Cad & "(" & Sql & ")" & ","
        
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
            Cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If


    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        If cadTABLA = "scafac" Then
            Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Else
            Sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        End If
        Sql = Sql & " VALUES " & Cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact = False
        cadErr = Err.Description
    Else
        InsertarLinFact = True
    End If
End Function





Private Function InsertarLinFact_new(cadTABLA As String, cadwhere As String, cadErr As String, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
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
            TipoFact = DevuelveValor("select codtipom from facturas where " & cadwhere)
            Select Case TipoFact
                Case "EAC" ' facturas a cuenta
                    '[Monica]21/03/2013: Estas facturas tambien se contabilizan sobr la cta de la variedad si la tienen
                    'CadCampo1 = vParamAplic.CtaVentasFraACta
                    CadCampo1 = "CASE tipomer.tiptimer WHEN 0 THEN ctavtasinterior WHEN 1 THEN ctavtasexportacion WHEN 2 THEN ctavtasindustria WHEN 3 THEN ctavtasretirada WHEN 4 THEN ctavtasotros END"
                    
                    CadCampo3 = "if(facturas.codvarie is null or facturas.codtimer is null," & DBSet(vParamAplic.CtaVentasFraACta, "T") & "," & CadCampo1 & ")"
                    
                    If vEmpresa.TieneAnalitica Then
                        Sql = " SELECT stipom.letraser,facturas.codtipom,numfactu,fecfactu," & CadCampo3 & " as cuenta,(baseimp1) as importe, '" & vParamAplic.CCosteFraACta & "' as codccost "
                    Else
                        Sql = " SELECT stipom.letraser,facturas.codtipom,numfactu,fecfactu," & CadCampo3 & " as cuenta,(baseimp1) as importe "
                    End If
                    
                    Sql = Sql & " FROM (facturas inner join usuarios.stipom stipom on facturas.codtipom=stipom.codtipom "
                    Sql = Sql & " LEFT JOIN variedades on facturas.codvarie = variedades.codvarie) "
                    Sql = Sql & " LEFT JOIN tipomer on facturas.codtimer = tipomer.codtimer "
                    
                    Sql = Sql & " WHERE " & cadwhere
                    If vEmpresa.TieneAnalitica Then
                        Sql = Sql & " GROUP BY 5,7 " '& cadCampo, codccost
                    Else
                        Sql = Sql & " GROUP BY 5 " '& cadCampo
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
                        Sql = " SELECT stipom.letraser,facturas_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost "
                    Else
                        Sql = " SELECT stipom.letraser,facturas_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
                    End If
                    
                    Sql = Sql & " FROM ((facturas_envases inner join usuarios.stipom stipom on facturas_envases.codtipom=stipom.codtipom) "
                    Sql = Sql & " inner join sartic on facturas_envases.codartic=sartic.codartic) "
                    Sql = Sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
                    Sql = Sql & " WHERE " & Replace(cadwhere, "facturas", "facturas_envases")
                    If vEmpresa.TieneAnalitica Then
                        Sql = Sql & " GROUP BY 5,7 " '& cadCampo, codccost
                    Else
                        Sql = Sql & " GROUP BY 5 " '& cadCampo
                    End If
                    '[Monica]05/05/2015: la suma debe ser distinta de 0
                    Sql = Sql & " HAVING sum(importel) <> 0 "
                    
                    Sql = Sql & "Union "
                    
                    ' LINEAS DE VARIEDADES
                    
                    If vEmpresa.TieneAnalitica Then
                        Sql = Sql & " SELECT stipom.letraser,facturas_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe, variedades.codccost "
                    Else
                        Sql = Sql & " SELECT stipom.letraser,facturas_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe "
                    End If
                    Sql = Sql & " FROM (((((facturas_variedad inner join usuarios.stipom stipom on facturas_variedad.codtipom=stipom.codtipom) "
                    Sql = Sql & " inner join albaran on facturas_variedad.numalbar = albaran.numalbar) "
                    Sql = Sql & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
                    Sql = Sql & " inner join albaran_variedad on facturas_variedad.numalbar = albaran_variedad.numalbar and facturas_variedad.numlinealbar = albaran_variedad.numlinea) "
                    Sql = Sql & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
                    Sql = Sql & " WHERE " & Replace(cadwhere, "facturas", "facturas_variedad")
                    If vEmpresa.TieneAnalitica Then
                        Sql = Sql & " GROUP BY 5,7 " '& cadCampo1, codccost
                    Else
                        Sql = Sql & " GROUP BY 5 " '& cadCampo1
                    End If
                    '[Monica]05/05/2015: la suma debe ser distinta de 0
                    Sql = Sql & " HAVING sum(impornet) <> 0 "
                    
                    ' LINEAS DE FACTURAS A CUENTA
                    Sql = Sql & "Union "
                    
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
                        Sql = Sql & " SELECT stipom.letraser,facturas_acuenta.codtipom,facturas_acuenta.numfactu,facturas_acuenta.fecfactu," & CadCampo3 & " as cuenta,(sum(baseimp1) * (-1)) as importe, variedades.codccost as codccost "
                    Else
                        Sql = Sql & " SELECT stipom.letraser,facturas_acuenta.codtipom,facturas_acuenta.numfactu,facturas_acuenta.fecfactu," & CadCampo3 & " as cuenta,(sum(baseimp1) * (-1)) as importe "
                    End If
                    Sql = Sql & " FROM (((facturas INNER JOIN facturas_acuenta ON facturas_acuenta.codtipomcta = facturas.codtipom and facturas_acuenta.numfactucta = facturas.numfactu and facturas_acuenta.fecfactucta = facturas.fecfactu) "
                    Sql = Sql & " LEFT JOIN variedades ON facturas.codvarie = variedades.codvarie) "
                    Sql = Sql & " LEFT JOIN tipomer ON facturas.codtimer = tipomer.codtimer) "
                    Sql = Sql & " INNER JOIN usuarios.stipom stipom ON facturas_acuenta.codtipom=stipom.codtipom"
                    Sql = Sql & " WHERE " & Replace(Replace(Replace(cadwhere, "facturas", "facturas_acuenta"), "numfactu", "facturas_acuenta.numfactu"), "fecfactu", "facturas_acuenta.fecfactu")
                    
                    
                    
                    If vEmpresa.TieneAnalitica Then
                        Sql = Sql & " GROUP BY 5,7 " '& cadCampo1, codccost
                    Else
                        Sql = Sql & " GROUP BY 5 " '& cadCampo1
                    End If
                    '[Monica]05/05/2015: la suma debe ser distinta de 0
                    Sql = Sql & " HAVING (sum(baseimp1) * (-1)) <> 0 "
    
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
                Sql = " SELECT stipom.letraser,facturassocio_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost "
            Else
                Sql = " SELECT stipom.letraser,facturassocio_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
            End If
            
            Sql = Sql & " FROM ((facturassocio_envases inner join usuarios.stipom stipom on facturassocio_envases.codtipom=stipom.codtipom) "
            Sql = Sql & " inner join sartic on facturassocio_envases.codartic=sartic.codartic) "
            Sql = Sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
            Sql = Sql & " WHERE " & Replace(cadwhere, "facturassocio", "facturassocio_envases")
            If vEmpresa.TieneAnalitica Then
                Sql = Sql & " GROUP BY 5,7 " '& cadCampo, codccost
            Else
                Sql = Sql & " GROUP BY 5 " '& cadCampo
            End If
            Sql = Sql & "Union "
            
            ' LINEAS DE VARIEDADES
            
            If vEmpresa.TieneAnalitica Then
                Sql = Sql & " SELECT stipom.letraser,facturassocio_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe, variedades.codccost "
            Else
                Sql = Sql & " SELECT stipom.letraser,facturassocio_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe "
            End If
            Sql = Sql & " FROM (((((facturassocio_variedad inner join usuarios.stipom stipom on facturassocio_variedad.codtipom=stipom.codtipom) "
            Sql = Sql & " inner join albaran on facturassocio_variedad.numalbar = albaran.numalbar) "
            Sql = Sql & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
            Sql = Sql & " inner join albaran_variedad on facturassocio_variedad.numalbar = albaran_variedad.numalbar and facturassocio_variedad.numlinealbar = albaran_variedad.numlinea) "
            Sql = Sql & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
            Sql = Sql & " WHERE " & Replace(cadwhere, "facturassocio", "facturassocio_variedad")
            If vEmpresa.TieneAnalitica Then
                Sql = Sql & " GROUP BY 5,7 " '& cadCampo1, codccost
            Else
                Sql = Sql & " GROUP BY 5 " '& cadCampo1
            End If
            
        
        Case "scafpc" 'COMPRAS
            'utilizamos sfamia.ctaventa o sfamia.aboventa
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctacompr"
            Else
                cadCampo = "sfamia.abocompr"
            End If
            If vEmpresa.TieneAnalitica Then
                Sql = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost"
            Else
                Sql = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
            End If
            Sql = Sql & " FROM (slifpc  "
            Sql = Sql & " inner join sartic on slifpc.codartic=sartic.codartic) "
            Sql = Sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
            Sql = Sql & " WHERE " & Replace(cadwhere, "scafpc", "slifpc")
            Sql = Sql & " GROUP BY " & cadCampo
            If vEmpresa.TieneAnalitica Then
                Sql = Sql & ", sfamia.codccost "
            End If
            
        Case "facturascom" 'facturas de compras de mercancias
            'utilizamos sfamia.ctaventa o sfamia.aboventa
            cadCampo = "variedades.ctacomtercero"
            If vEmpresa.TieneAnalitica Then
                Sql = " SELECT facturascom_variedad.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importe) as importe, variedades.codccost"
            Else
                Sql = " SELECT facturascom_variedad.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importe) as importe"
            End If
            Sql = Sql & " FROM (facturascom_variedad  "
            Sql = Sql & " inner join variedades on facturascom_variedad.codvarie=variedades.codvarie) "
            Sql = Sql & " WHERE " & Replace(cadwhere, "facturascom", "facturascom_variedad")
            Sql = Sql & " GROUP BY " & cadCampo
            If vEmpresa.TieneAnalitica Then
                Sql = Sql & ", variedades.codccost "
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
            Sql = "select tipo from tcafpc where " & cadwhere
            Tipo = DevuelveValor(Sql) ' 0=transportista 1=comisionista
            Select Case Tipo
                Case 0 ' Transportista
        '++monica: si tipomercado = 1(exportacion) cogemos  variedades.ctatraexporta
        '          si tipomercado <> 1 (distinto de exportacion) cogemos  variedades.ctatrainterior
                    If vEmpresa.TieneAnalitica Then
                         Sql = " SELECT 1, if(tipomer.tiptimer = 1,variedades.ctatraexporta,variedades.ctatrainterior) as cuenta, sum(tlifpc.importel) as importe, variedades.codccost "
                    Else
                         Sql = " SELECT 1, if(tipomer.tiptimer = 1,variedades.ctatraexporta,variedades.ctatrainterior) as cuenta, sum(tlifpc.importel) as importe, '----' "
                    End If
                    Sql = Sql & " FROM tlifpc, albaran, albaran_variedad, variedades, tipomer"
                    Sql = Sql & " WHERE " & Replace(cadwhere, "tcafpc", "tlifpc") & " and"
                    Sql = Sql & " tlifpc.numalbar = albaran_variedad.numalbar and "
                    Sql = Sql & " tlifpc.numlinea = albaran_variedad.numlinea and "
                    Sql = Sql & " albaran_variedad.numalbar = albaran.numalbar and "
                    Sql = Sql & " albaran_variedad.codvarie = variedades.codvarie and "
                    Sql = Sql & " albaran.codtimer = tipomer.codtimer "
                    Sql = Sql & " group by 1,2 "
                    Sql = Sql & " union "
                    Sql = Sql & " select 2, codmacta as cuenta, importel as importe, '----' "
                    Sql = Sql & " from tcafpv "
                    Sql = Sql & " where " & Replace(cadwhere, "tcafpc", "tcafpv")
                    Sql = Sql & " group by 1,2 "
                    Sql = Sql & " order by 1,2 "

                Case 1 ' Comisionista
                    If vEmpresa.TieneAnalitica Then
                         Sql = " SELECT 1, variedades.ctacomisionista as cuenta, sum(tlifpc.importel) as importe, variedades.codccost "
                    Else
                         Sql = " SELECT 1, variedades.ctacomisionista as cuenta, sum(tlifpc.importel) as importe, '----' "
                    End If
                    Sql = Sql & " FROM tlifpc, albaran, albaran_variedad, variedades "
                    Sql = Sql & " WHERE " & Replace(cadwhere, "tcafpc", "tlifpc") & " and"
                    Sql = Sql & " tlifpc.numalbar = albaran_variedad.numalbar and "
                    Sql = Sql & " tlifpc.numlinea = albaran_variedad.numlinea and "
                    Sql = Sql & " albaran_variedad.numalbar = albaran.numalbar and "
                    Sql = Sql & " albaran_variedad.codvarie = variedades.codvarie "
                    Sql = Sql & " group by 1,2 "
                    Sql = Sql & " order by 1,2 "
            End Select
    End Select
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = Cad
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
        Sql = ""
        Sql2 = ""
        
        If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then 'VENTAS a clientes
            Sql = "'" & Rs!letraser & "'," & Rs!NumFactu & "," & Year(Rs!FecFactu) & "," & i & ","
            Sql = Sql & DBSet(Rs!Cuenta, "T")
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
            If cadTABLA = "scafpc" Or cadTABLA = "facturascom" Then 'COMPRAS
                'Laura 24/10/2006
                'SQL = numRegis & "," & Year(RS!FecFactu) & "," & i & ","
                Sql = numRegis & "," & AnyoFacPr & "," & i & ","
                
    '            If ImpLinea >= 0 Then
                    Sql = Sql & DBSet(Rs!Cuenta, "T")
    '            Else
    '                SQL = SQL & DBSet(RS!abocompr, "T")
    '            End If
            Else 'TRANSPORTE
                Sql = numRegis & "," & AnyoFacPr & "," & i & ","
                Sql = Sql & DBSet(Rs!Cuenta, "T")
            End If
        End If
        
        Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la �ltima hay q descontarle para q coincida con total factura
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
            If cadTABLA = "tcafpc" Then
                If DBLet(Rs!CodCCost, "T") = "----" Then
                    Sql = Sql & DBSet(CCoste, "T")
                Else
                    Sql = Sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBLet(Rs!CodCCost, "T")
                End If
            Else
                Sql = Sql & DBSet(Rs!CodCCost, "T")
                CCoste = DBLet(Rs!CodCCost, "T")
            End If
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        Cad = Cad & "(" & Sql & ")" & ","
        
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
            Cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If


    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
            Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Else
            Sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        End If
        Sql = Sql & " VALUES " & Cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_new = False
        cadErr = Err.Description
    Else
        InsertarLinFact_new = True
    End If
End Function


Private Function InsertarLinFact_newContaNueva(cadTABLA As String, cadwhere As String, cadErr As String, Optional numRegis As Long, Optional FraIntraCom As String, Optional FRecep As String, Optional TipoIvaFra As Byte) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
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

Dim IvaABuscar As Integer

    On Error GoTo EInLinea
    

    Select Case cadTABLA
        Case "facturas" 'VENTAS
            '[Monica]23/06/2010 Miramos que tipo de movimiento tiene la factura
            ' si es EAC (factura de anticipo) la cuenta base est� en parametros
            TipoFact = DevuelveValor("select codtipom from facturas where " & cadwhere)
            Select Case TipoFact
                Case "EAC" ' facturas a cuenta
                    '[Monica]21/03/2013: Estas facturas tambien se contabilizan sobr la cta de la variedad si la tienen
                    'CadCampo1 = vParamAplic.CtaVentasFraACta
                    CadCampo1 = "CASE tipomer.tiptimer WHEN 0 THEN ctavtasinterior WHEN 1 THEN ctavtasexportacion WHEN 2 THEN ctavtasindustria WHEN 3 THEN ctavtasretirada WHEN 4 THEN ctavtasotros END"
                    
                    CadCampo3 = "if(facturas.codvarie is null or facturas.codtimer is null," & DBSet(vParamAplic.CtaVentasFraACta, "T") & "," & CadCampo1 & ")"
                    
                    If vEmpresa.TieneAnalitica Then
                        Sql = " SELECT stipom.letraser,facturas.codtipom,numfactu,fecfactu," & CadCampo3 & " as cuenta,(baseimp1) as importe, '" & vParamAplic.CCosteFraACta & "' as codccost "
                    Else
                        Sql = " SELECT stipom.letraser,facturas.codtipom,numfactu,fecfactu," & CadCampo3 & " as cuenta,(baseimp1) as importe "
                    End If
                    
                    Sql = Sql & ",facturas.codiiva1 codigiva, facturas.porciva1 porciva, facturas.porcrec1 porcrec "
                    
                    Sql = Sql & " FROM (facturas inner join usuarios.stipom stipom on facturas.codtipom=stipom.codtipom "
                    Sql = Sql & " LEFT JOIN variedades on facturas.codvarie = variedades.codvarie) "
                    Sql = Sql & " LEFT JOIN tipomer on facturas.codtimer = tipomer.codtimer "
                    
                    Sql = Sql & " WHERE " & cadwhere
                    If vEmpresa.TieneAnalitica Then
                        Sql = Sql & " GROUP BY 5,7 " '& cadCampo, codccost
                    Else
                        Sql = Sql & " GROUP BY 5 " '& cadCampo
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
                        Sql = " SELECT stipom.letraser,facturas_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost "
                    Else
                        Sql = " SELECT stipom.letraser,facturas_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
                    End If
                    
                    Sql = Sql & ", facturas_envases.codigiva, tiposiva.porceiva porciva,  tiposiva.porcerec porcrec"
                    
                    Sql = Sql & " FROM (((facturas_envases inner join usuarios.stipom stipom on facturas_envases.codtipom=stipom.codtipom) "
                    Sql = Sql & " inner join sartic on facturas_envases.codartic=sartic.codartic) "
                    Sql = Sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia) "
                    Sql = Sql & " inner join ariconta" & vParamAplic.NumeroConta & ".tiposiva on ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = facturas_envases.codigiva"
                    Sql = Sql & " WHERE " & Replace(cadwhere, "facturas", "facturas_envases")
                    If vEmpresa.TieneAnalitica Then
                        Sql = Sql & " GROUP BY 5,7,8 " '& cadCampo, codccost
                    Else
                        Sql = Sql & " GROUP BY 5,7 " '& cadCampo
                    End If
                    '[Monica]05/05/2015: la suma debe ser distinta de 0
                    Sql = Sql & " HAVING sum(importel) <> 0 "
                    
                    Sql = Sql & "Union "
                    
                    ' LINEAS DE VARIEDADES
                    
                    If vEmpresa.TieneAnalitica Then
                        Sql = Sql & " SELECT stipom.letraser,facturas_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe, variedades.codccost "
                    Else
                        Sql = Sql & " SELECT stipom.letraser,facturas_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe "
                    End If
                    Sql = Sql & ", facturas_variedad.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                    
                    Sql = Sql & " FROM (((((facturas_variedad inner join usuarios.stipom stipom on facturas_variedad.codtipom=stipom.codtipom) "
                    Sql = Sql & " inner join albaran on facturas_variedad.numalbar = albaran.numalbar) "
                    Sql = Sql & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
                    Sql = Sql & " inner join albaran_variedad on facturas_variedad.numalbar = albaran_variedad.numalbar and facturas_variedad.numlinealbar = albaran_variedad.numlinea) "
                    Sql = Sql & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
                    Sql = Sql & " inner join ariconta" & vParamAplic.NumeroConta & ".tiposiva on ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = facturas_variedad.codigiva "
                    Sql = Sql & " WHERE " & Replace(cadwhere, "facturas", "facturas_variedad")
                    
                    If vEmpresa.TieneAnalitica Then
                        Sql = Sql & " GROUP BY 5,7,8 " '& cadCampo1, codccost, codigiva
                    Else
                        Sql = Sql & " GROUP BY 5,7  " '& cadCampo1, codigiva
                    End If
                    '[Monica]05/05/2015: la suma debe ser distinta de 0
                    Sql = Sql & " HAVING sum(impornet) <> 0 "
                    
                    ' LINEAS DE FACTURAS A CUENTA
                    Sql = Sql & "Union "
                    
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
                        Sql = Sql & " SELECT stipom.letraser,facturas_acuenta.codtipom,facturas_acuenta.numfactu,facturas_acuenta.fecfactu," & CadCampo3 & " as cuenta,(sum(baseimp1) * (-1)) as importe, variedades.codccost as codccost "
                    Else
                        Sql = Sql & " SELECT stipom.letraser,facturas_acuenta.codtipom,facturas_acuenta.numfactu,facturas_acuenta.fecfactu," & CadCampo3 & " as cuenta,(sum(baseimp1) * (-1)) as importe "
                    End If
                    Sql = Sql & ", facturas.codiiva1 codigiva, facturas.porciva1 porciva, facturas.porcrec1 porcrec "
                    
                    Sql = Sql & " FROM (((facturas INNER JOIN facturas_acuenta ON facturas_acuenta.codtipomcta = facturas.codtipom and facturas_acuenta.numfactucta = facturas.numfactu and facturas_acuenta.fecfactucta = facturas.fecfactu) "
                    Sql = Sql & " LEFT JOIN variedades ON facturas.codvarie = variedades.codvarie) "
                    Sql = Sql & " LEFT JOIN tipomer ON facturas.codtimer = tipomer.codtimer) "
                    Sql = Sql & " INNER JOIN usuarios.stipom stipom ON facturas_acuenta.codtipom=stipom.codtipom"
                    Sql = Sql & " WHERE " & Replace(Replace(Replace(cadwhere, "facturas", "facturas_acuenta"), "numfactu", "facturas_acuenta.numfactu"), "fecfactu", "facturas_acuenta.fecfactu")
                    
                    If vEmpresa.TieneAnalitica Then
                        Sql = Sql & " GROUP BY 5,7,8 " '& cadCampo1, codccost, codigiva
                    Else
                        Sql = Sql & " GROUP BY 5,7 " '& cadCampo1, codigiva
                        
                    End If
                    
                    '[Monica]05/05/2015: la suma debe ser distinta de 0
                    Sql = Sql & " HAVING (sum(baseimp1) * (-1)) <> 0 "
    
                    If vEmpresa.TieneAnalitica Then
                        Sql = Sql & " ORDER BY 8,5,7 "
                    Else
                        Sql = Sql & " ORDER BY 7,5 "
                    End If
    
    
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
                Sql = " SELECT stipom.letraser,facturassocio_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost "
            Else
                Sql = " SELECT stipom.letraser,facturassocio_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
            End If
            
            Sql = Sql & ", facturassocio_envases.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
            
            Sql = Sql & " FROM (((facturassocio_envases inner join usuarios.stipom stipom on facturassocio_envases.codtipom=stipom.codtipom) "
            Sql = Sql & " inner join sartic on facturassocio_envases.codartic=sartic.codartic) "
            Sql = Sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia) "
            Sql = Sql & " inner join ariconta" & vParamAplic.NumeroConta & ".tiposiva on ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = facturassocio_envases.codigiva "
            Sql = Sql & " WHERE " & Replace(cadwhere, "facturassocio", "facturassocio_envases")
            If vEmpresa.TieneAnalitica Then
                Sql = Sql & " GROUP BY 5,7,8 " '& cadCampo, codccost
            Else
                Sql = Sql & " GROUP BY 5,7 " '& cadCampo
            End If
            Sql = Sql & "Union "
            
            ' LINEAS DE VARIEDADES
            
            If vEmpresa.TieneAnalitica Then
                Sql = Sql & " SELECT stipom.letraser,facturassocio_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe, variedades.codccost "
            Else
                Sql = Sql & " SELECT stipom.letraser,facturassocio_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe "
            End If
            Sql = Sql & ", facturassocio_variedad.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
            
            Sql = Sql & " FROM (((((facturassocio_variedad inner join usuarios.stipom stipom on facturassocio_variedad.codtipom=stipom.codtipom) "
            Sql = Sql & " inner join albaran on facturassocio_variedad.numalbar = albaran.numalbar) "
            Sql = Sql & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
            Sql = Sql & " inner join albaran_variedad on facturassocio_variedad.numalbar = albaran_variedad.numalbar and facturassocio_variedad.numlinealbar = albaran_variedad.numlinea) "
            Sql = Sql & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
            Sql = Sql & " inner join ariconta" & vParamAplic.NumeroConta & ".tiposiva on ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = facturassocio_variedad.codigiva "
            
            Sql = Sql & " WHERE " & Replace(cadwhere, "facturassocio", "facturassocio_variedad")
            If vEmpresa.TieneAnalitica Then
                Sql = Sql & " GROUP BY 5,7,8 " '& cadCampo1, codccost, codigiva
                Sql = Sql & " ORDER BY 8,5,7 "
            Else
                Sql = Sql & " GROUP BY 5,7 " '& cadCampo1, codigiva
                Sql = Sql & " ORDER BY 7,5 "
                
            End If
            
        
        Case "scafpc" 'COMPRAS
            'utilizamos sfamia.ctaventa o sfamia.aboventa
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctacompr"
            Else
                cadCampo = "sfamia.abocompr"
            End If
            If vEmpresa.TieneAnalitica Then
                Sql = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost"
                If FraIntraCom <> "" Then
                    Sql = Sql & "," & DBSet(FraIntraCom, "N") & ",sartic.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                Else
                    Sql = Sql & ",sartic.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                End If
            Else
                Sql = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
                If FraIntraCom <> "" Then
                    Sql = Sql & "," & DBSet(FraIntraCom, "N") & ",sartic.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                Else
                    Sql = Sql & ",sartic.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                End If
            End If
            Sql = Sql & " FROM ((slifpc  "
            Sql = Sql & " inner join sartic on slifpc.codartic=sartic.codartic) "
            Sql = Sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia) "
            Sql = Sql & " inner join ariconta" & vParamAplic.NumeroConta & ".tiposiva on ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = sartic.codigiva "
            Sql = Sql & " WHERE " & Replace(cadwhere, "scafpc", "slifpc")
            Sql = Sql & " GROUP BY " & cadCampo & ", codigiva "
            If vEmpresa.TieneAnalitica Then
                Sql = Sql & ", sfamia.codccost "
            End If
        
        Case "facturascom" 'COMPRAS DE VARIEDADES
            'utilizamos sfamia.ctaventa o sfamia.aboventa
            cadCampo = "variedades.ctacomtercero"
            If vEmpresa.TieneAnalitica Then
                Sql = " SELECT facturascom_variedad.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importe) as importe, variedades.codccost "
                If FraIntraCom <> "" Then
                    Sql = Sql & "," & DBSet(FraIntraCom, "N") & ",facturascom_variedad.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                Else
                    Sql = Sql & ",facturascom_variedad.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                End If
            Else
                Sql = " SELECT facturascom_variedad.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importe) as importe"
                If FraIntraCom <> "" Then
                    Sql = Sql & "," & DBSet(FraIntraCom, "N") & ",facturascom_variedad.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                Else
                    Sql = Sql & ",facturascom_variedad.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                End If
            End If
            Sql = Sql & " FROM (facturascom_variedad  "
            Sql = Sql & " inner join variedades on facturascom_variedad.codvarie=variedades.codvarie) "
            Sql = Sql & " inner join ariconta" & vParamAplic.NumeroConta & ".tiposiva on ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = facturascom_variedad.codigiva "
            Sql = Sql & " WHERE " & Replace(cadwhere, "facturascom", "facturascom_variedad")
            Sql = Sql & " GROUP BY " & cadCampo & ", codigiva "
            If vEmpresa.TieneAnalitica Then
                Sql = Sql & ", variedades.codccost "
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
            Sql = "select tipo from tcafpc where " & cadwhere
            Tipo = DevuelveValor(Sql) ' 0=transportista 1=comisionista
            
            '[Monica]17/07/2017: Iva1
            Dim Iva1 As Integer
            Sql = "select tipoiva1 from tcafpc where " & cadwhere
            Iva1 = DevuelveValor(Sql) ' 0=transportista 1=comisionista
            
            
            Select Case Tipo
                Case 0 ' Transportista
        '++monica: si tipomercado = 1(exportacion) cogemos  variedades.ctatraexporta
        '          si tipomercado <> 1 (distinto de exportacion) cogemos  variedades.ctatrainterior
       
                    If vEmpresa.TieneAnalitica Then
                         Sql = " SELECT 1, if(tipomer.tiptimer = 1,variedades.ctatraexporta,variedades.ctatrainterior) as cuenta, sum(tlifpc.importel) as importe, variedades.codccost "
                         Sql = Sql & ", " & Iva1 & " codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                    Else
                         Sql = " SELECT 1, if(tipomer.tiptimer = 1,variedades.ctatraexporta,variedades.ctatrainterior) as cuenta, sum(tlifpc.importel) as importe, '----' "
                         Sql = Sql & ", " & Iva1 & " codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec"
                    End If
                    Sql = Sql & " FROM tlifpc, albaran, albaran_variedad, variedades, tipomer, ariconta" & vParamAplic.NumeroConta & ".tiposiva "
                    Sql = Sql & " WHERE " & Replace(cadwhere, "tcafpc", "tlifpc") & " and"
                    Sql = Sql & " tlifpc.numalbar = albaran_variedad.numalbar and "
                    Sql = Sql & " tlifpc.numlinea = albaran_variedad.numlinea and "
                    Sql = Sql & " albaran_variedad.numalbar = albaran.numalbar and "
                    Sql = Sql & " albaran_variedad.codvarie = variedades.codvarie and "
                    Sql = Sql & " albaran.codtimer = tipomer.codtimer "
                    Sql = Sql & " and ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = " & DBSet(Iva1, "N")
                    Sql = Sql & " group by 1,2,4,5,6,7 "
                    Sql = Sql & " union "
                    Sql = Sql & " select 2, codmacta as cuenta, importel as importe, '----' "
                    Sql = Sql & ", " & Iva1 & " codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                    Sql = Sql & " from tcafpv, ariconta" & vParamAplic.NumeroConta & ".tiposiva "
                    Sql = Sql & " where " & Replace(cadwhere, "tcafpc", "tcafpv")
                    Sql = Sql & " and ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = " & DBSet(Iva1, "N")
                    Sql = Sql & " group by 1,2,4,5,6,7 "
                    Sql = Sql & " order by 1,2 "

                Case 1 ' Comisionista
                    If vEmpresa.TieneAnalitica Then
                         Sql = " SELECT 1, variedades.ctacomisionista as cuenta, sum(tlifpc.importel) as importe, variedades.codccost "
                         Sql = Sql & ", " & Iva1 & " codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                    Else
                         Sql = " SELECT 1, variedades.ctacomisionista as cuenta, sum(tlifpc.importel) as importe, '----' "
                         Sql = Sql & ", " & Iva1 & " codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec "
                    End If
                    Sql = Sql & " FROM tlifpc, albaran, albaran_variedad, variedades, ariconta" & vParamAplic.NumeroConta & ".tiposiva "
                    Sql = Sql & " WHERE " & Replace(cadwhere, "tcafpc", "tlifpc") & " and"
                    Sql = Sql & " tlifpc.numalbar = albaran_variedad.numalbar and "
                    Sql = Sql & " tlifpc.numlinea = albaran_variedad.numlinea and "
                    Sql = Sql & " albaran_variedad.numalbar = albaran.numalbar and "
                    Sql = Sql & " albaran_variedad.codvarie = variedades.codvarie "
                    Sql = Sql & " and ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = " & DBSet(Iva1, "N")
                    Sql = Sql & " group by 1,2,4,5,6,7 "
                    Sql = Sql & " order by 1,2 "
            End Select
            
            Dim Sql7 As String
            Sql7 = "select fecrecep from tcafpc where " & cadwhere
            FRecep = DevuelveValor(Sql7) '
    End Select
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenKeyset, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    totimp = 0
    SQLaux = ""
    
    
    While Not Rs.EOF
        SQLaux = Cad
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
        Sql = ""
        Sql2 = ""
        
        If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then 'VENTAS a clientes
            Sql = "'" & Rs!letraser & "'," & Rs!NumFactu & "," & Year(Rs!FecFactu) & "," & i & ","
            Sql = Sql & DBSet(Rs!Cuenta, "T") & ","
        Else
            If cadTABLA = "scafpc" Or cadTABLA = "facturascom" Then 'COMPRAS
                'Laura 24/10/2006
                'SQL = numRegis & "," & Year(RS!FecFactu) & "," & i & ","
                Sql = DBSet(SerieFraPro, "T") & "," & numRegis & "," & DBSet(FRecep, "F") & "," & AnyoFacPr & "," & i & ","
                Sql = Sql & DBSet(Rs!Cuenta, "T") & ","
            
            Else 'TRANSPORTE
                Sql = DBSet(SerieFraPro, "T") & "," & numRegis & "," & DBSet(FRecep, "F") & "," & AnyoFacPr & "," & i & ","
                Sql = Sql & DBSet(Rs!Cuenta, "T") & ","
            End If
        End If
        
        'Vemos que tipo de IVA es en el vector de importes
'        NumeroIVA = 127
'        For k = 0 To 2
'            If Rs!Codigiva = vTipoIva(k) Then
'                NumeroIVA = k
'                Exit For
'            End If
'        Next
'        If NumeroIVA > 100 Then Err.Raise 513, "Error obteniendo IVA: " & Rs!Codigiva
'de David
        NumeroIVA = 127
        For k = 0 To 2
        
            IvaABuscar = Rs!Codigiva
            'JUNIO 18
            ''0 Normal   1 R.E     2 Exento
            If TipoIvaFra = 1 Then
'                If IvaABuscar = vParamAplic.TipoIVA1 Then IvaABuscar = vParamAplic.TipoIVAre1
'                If IvaABuscar = vParamAplic.TipoIVA2 Then IvaABuscar = vParamAplic.TipoIVAre2
'                If IvaABuscar = vParamAplic.TipoIVA3 Then IvaABuscar = vParamAplic.TipoIVAre3
                 IvaABuscar = vTipoIva(k)
            Else
                If TipoIvaFra = 2 Then
                    'Solo tiene un IVA
                    IvaABuscar = vTipoIva(k)
                    
                End If
            End If
            
        
        
            If IvaABuscar = vTipoIva(k) Then
            'If Rs!Codigiva = vTipoIva(K) Then
                NumeroIVA = k
                Exit For
            End If
        Next
        If NumeroIVA > 100 Then Err.Raise 513, , "Error obteniendo IVA: " & Rs!Codigiva
'        If PrimerCodigiva < 0 Then PrimerCodigiva = k
'hasta aqui
        
        If vEmpresa.TieneAnalitica Then
            If cadTABLA = "tcafpc" Then
                If DBLet(Rs!CodCCost, "T") = "----" Then
                    Sql = Sql & DBSet(CCoste, "T")
                Else
                    Sql = Sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBLet(Rs!CodCCost, "T")
                End If
            Else
                Sql = Sql & DBSet(Rs!CodCCost, "T")
                CCoste = DBLet(Rs!CodCCost, "T")
            End If
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
            Sql = Sql & "," & DBSet(Rs!FecFactu, "F")
        End If
        
        vBaseIva(NumeroIVA) = vBaseIva(NumeroIVA) - ImpLinea   'Para ajustar el importe y que no haya descuadre
        
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
        
        
        HayQueAjustar = False
        If vBaseIva(NumeroIVA) <> 0 Or vImpIva(NumeroIVA) <> 0 Or vImpRec(NumeroIVA) <> 0 Then
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
                If Rs!Codigiva <> vTipoIva(NumeroIVA) Then
                    'NO es el mismo tipo de IVA
                    'Hay que ajustar
                    HayQueAjustar = True
                End If
            End If
            Rs.MovePrevious
        End If
        
        Sql = Sql & "," & vTipoIva(NumeroIVA) & "," & DBSet(vPorcIva(NumeroIVA), "N") & "," & DBSet(vPorcRec(NumeroIVA), "N", "S") & ","
        
        If HayQueAjustar Then
            
            If vBaseIva(NumeroIVA) <> 0 Then ImpLinea = ImpLinea + vBaseIva(NumeroIVA)
            If vImpIva(NumeroIVA) <> 0 Then ImpImva = ImpImva + vImpIva(NumeroIVA)
            If vImpRec(NumeroIVA) <> 0 Then ImpREC = ImpREC + vImpRec(NumeroIVA)
            
        End If

        
'        'Caluclo el importe de IVA y el de recargo de equivalencia
'        ImpImva = vPorcIva(NumeroIVA) / 100
'        ImpImva = Round2(ImpLinea * ImpImva, 2)
'        If vPorcRec(NumeroIVA) = 0 Then
'            ImpREC = 0
'        Else
'            ImpREC = vPorcRec(NumeroIVA) / 100
'            ImpREC = Round2(ImpLinea * ImpREC, 2)
'        End If
'        vImpIva(NumeroIVA) = vImpIva(NumeroIVA) - ImpImva
'        vImpRec(NumeroIVA) = vImpRec(NumeroIVA) - ImpREC
        
        
        ' baseimpo , impoiva, imporec, aplicret, CodCCost
        Sql = Sql & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpREC, "N", "S")
        
        ' si la linea lleva retencion
        If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then 'VENTAS a clientes
        Else
            '[Monica]18/07/2017: lleve o no lleve retencion lo marcamos como que se le aplica
            Sql = Sql & ",1"
        End If
      
'        Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la �ltima hay q descontarle para q coincida con total factura
'        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        Cad = Cad & "(" & Sql & ")" & ","
        
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    

    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
            Sql = "INSERT INTO factcli_lineas (numserie,numfactu,anofactu,numlinea,codmacta,codccost,fecfactu,codigiva,porciva,porcrec,baseimpo,impoiva,imporec) "
        Else
            Sql = "INSERT INTO factpro_lineas (numserie,numregis,fecharec,anofactu,numlinea,codmacta,codccost,codigiva,porciva,porcrec,baseimpo,impoiva,imporec,aplicret) "
        End If
        Sql = Sql & " VALUES " & Cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_newContaNueva = False
        cadErr = Err.Description
    Else
        InsertarLinFact_newContaNueva = True
    End If
End Function



Private Function ActualizarCabFact(cadTABLA As String, cadwhere As String, cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    
    Sql = "UPDATE " & cadTABLA & " SET intconta=1 "
    Sql = Sql & " WHERE " & cadwhere

    conn.Execute Sql
    
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

Public Function PasarFacturaProv(cadwhere As String, CodCCost As String, FechaFin As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariges.scafpc --> conta.cabfactprov
' ariges.slifpc --> conta.linfactprov
'Actualizar la tabla ariges.scafpc.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim Sql As String
Dim Mc As Contadores
Dim FraIntraCom2 As String

Dim TipoIvaFactura As Byte '0 Normal   1 R.E     2 Exento    JULIO 18

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
        
    
    Set Mc = New Contadores
    
    '---- Insertar en la conta Cabecera Factura
    TipoIvaFactura = 0
    b = InsertarCabFactProv(cadwhere, cadMen, Mc, FechaFin, vContaFra, FraIntraCom2, TipoIvaFactura)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        CCoste = CodCCost
        '---- Insertar lineas de Factura en la Conta
        If Not vParamAplic.ContabilidadNueva Then
            b = InsertarLinFact_new("scafpc", cadwhere, cadMen, Mc.Contador)
        Else
            b = InsertarLinFact_newContaNueva("scafpc", cadwhere, cadMen, Mc.Contador, FraIntraCom2, , TipoIvaFactura)
        End If
        cadMen = "Insertando Lin. Factura: " & cadMen

        If b Then
            If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaProv(vContaFra.NumeroFactura, vContaFra.Anofac)
            
            '---- Poner intconta=1 en ariges.scafac
            b = ActualizarCabFact("scafpc", cadwhere, cadMen)
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
            InsertarTMPErrFac cadMen, cadwhere
'            SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
'            SQL = SQL & " Select *," & DBSet(Mid(cadMen, 1, 200), "T") & " as error From tmpFactu "
'            SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
'            Conn.Execute SQL
        End If
    End If
End Function


Private Function InsertarCabFactProv(cadwhere As String, cadErr As String, ByRef Mc As Contadores, FechaFin As String, ByRef vCF As cContabilizarFacturas, ByRef EsFacturaIntracom2 As String, ByRef QueTipoDeIVA As Byte) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el a�o de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Intracom As Integer

Dim TipoOpera As Byte
Dim CadenaInsertFaclin2     As String
Dim ImporAux As Currency

Dim Aux As String
Dim Sql2 As String

    On Error GoTo EInsertar
       
    
    Sql = " SELECT fecfactu,year(fecrecep) as anofacpr,fecrecep,numfactu,proveedor.codmacta,"
    Sql = Sql & "scafpc.dtoppago,scafpc.dtognral,baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    Sql = Sql & "totalfac,tipoiva1,tipoiva2,tipoiva3,proveedor.codprove, proveedor.nomprove, proveedor.tipprove, scafpc.confacpr "
    Sql = Sql & ",scafpc.domprove,scafpc.codpobla,scafpc.pobprove,scafpc.proprove,scafpc.nifprove,scafpc.codforpa "
    '[Monica]02/05/2019: incluimos si la factura es rectificativa
    Sql = Sql & ",scafpc.esrectificativa "
    Sql = Sql & " FROM " & "scafpc "
    Sql = Sql & "INNER JOIN " & "proveedor ON scafpc.codprove=proveedor.codprove "
    Sql = Sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
    
        If Mc.ConseguirContador("1", (Rs!FecRecep <= CDate(FechaFin) - 365), True) = 0 Then
        
            vCF.NumeroFactura = Mc.Contador
            vCF.Anofac = Year(DBLet(Rs!FecRecep))
            vvFecRecep = DBLet(Rs!FecRecep)
        
            '[Monica]27/07/2017
            vCF.observa = DBLet(Rs!confacpr)
        
            
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
            '[Monica]09/06/2017: antes era baseiva2 y baseiva3
            If DBLet(Rs!TipoIVA2, "N") = "0" Then Nulo2 = "S"
            If DBLet(Rs!TipoIVA3, "N") = "0" Then Nulo3 = "S"
            Sql = ""
            If vParamAplic.ContabilidadNueva Then Sql = "'" & SerieFraPro & "',"
            
            Sql = Sql & Mc.Contador & "," & DBSet(Rs!FecFactu, "F") & "," & Rs!anofacpr & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!NumFactu, "T") & "," & DBSet(Rs!Codmacta, "T") & "," & DBSet(Rs!confacpr, "T") & ","
            
            If Not vParamAplic.ContabilidadNueva Then
                Sql = Sql & DBSet(Rs!BaseIVA1, "N") & "," & DBSet(Rs!BaseIVA2, "N", "S") & "," & DBSet(Rs!BaseIVA3, "N", "S") & ","
                Sql = Sql & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", Nulo2) & "," & DBSet(Rs!porciva3, "N", Nulo3) & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & DBSet(Rs!impoiva2, "N", Nulo2) & "," & DBSet(Rs!impoiva3, "N", Nulo3) & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!TipoIVA1, "N") & "," & DBSet(Rs!TipoIVA2, "N", Nulo2) & "," & DBSet(Rs!TipoIVA3, "N", Nulo3) & "," & DBSet(Intracom, "N") & ","
            Else
                Sql = Sql & DBSet(Rs!nomprove, "T") & "," & DBSet(Rs!domprove, "T", "S") & ","
                Sql = Sql & DBSet(Rs!codPobla, "T", "S") & "," & DBSet(Rs!pobprove, "T", "S") & "," & DBSet(Rs!proprove, "T", "S") & ","
                Sql = Sql & DBSet(Rs!nifProve, "F", "S") & ",'ES',"
                Sql = Sql & DBSet(Rs!Codforpa, "N") & ","
                
                TipoOpera = 0
                QueTipoDeIVA = 0
                 'IVA ES CERO
                If Rs!tipprove = 1 Then
                    'intracomunitaria
                    TipoOpera = 1
                    QueTipoDeIVA = 2
                Else
                    'Exstranjero
                     If Rs!tipprove = 1 Then
                        TipoOpera = 2
                        QueTipoDeIVA = 2
                     End If
                End If
                
                Aux = "0"
                Select Case TipoOpera
                Case 0
'[Monica]08/06/2017: no s� cuando es rectificativa en proveedores
'                    If Rs!TotalFac < 0 Then
'                        Aux = "D"
'                    Else

                     '[Monica]02/05/2019: a�adimos la columna de si es rectificativa
                     If DBLet(Rs!EsRectificativa, "N") = 1 Then
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
                Sql = Sql & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                
                
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
                vImpIva(0) = DBLet(Rs!impoiva1, "N")
                vImpRec(0) = 0
                vBaseIva(0) = DBLet(Rs!BaseIVA1, "N")
                
                vTipoIva(1) = 0: vTipoIva(2) = 0
                
                If Not IsNull(Rs!porciva2) Then
                    Sql2 = Aux & "2," & DBSet(Rs!BaseIVA2, "N") & "," & Rs!TipoIVA2 & "," & DBSet(Rs!porciva2, "N") & ","
                    Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva2, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                    vTipoIva(1) = Rs!TipoIVA2
                    vPorcIva(1) = Rs!porciva2
                    vPorcRec(1) = 0
                    vImpIva(1) = DBLet(Rs!impoiva2, "N")
                    vImpRec(1) = 0
                    vBaseIva(1) = DBLet(Rs!BaseIVA2, "N")
                
                End If
                If Not IsNull(Rs!porciva3) Then
                    Sql2 = Aux & "3," & DBSet(Rs!BaseIVA3, "N") & "," & Rs!TipoIVA3 & "," & DBSet(Rs!porciva3, "N") & ","
                    Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva3, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                    vTipoIva(2) = Rs!TipoIVA3
                    vPorcIva(2) = Rs!porciva3
                    vPorcRec(2) = 0
                    vImpIva(2) = DBLet(Rs!impoiva3, "N")
                    vImpRec(2) = 0
                    vBaseIva(2) = DBLet(Rs!BaseIVA3, "N")
                End If
                
                    
                    
                'Los totales
                'totbases,totbasesret,totivas,totrecargo,totfacpr,
                ImporAux = Rs!BaseIVA1 + DBLet(Rs!BaseIVA2, "N") + DBLet(Rs!BaseIVA3, "N")
                Sql = Sql & DBSet(ImporAux, "N") & "," & ValorNulo & ","
                'totivas
                ImporAux = Rs!impoiva1 + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
                Sql = Sql & DBSet(ImporAux, "N") & "," & DBSet(Rs!TotalFac, "N") & ","
                        
                  
                EsFacturaIntracom2 = ""
                If DBLet(Rs!tipprove, "N") = 1 Then
                    'OK es intracomunitaria
                    EsFacturaIntracom2 = Rs!TipoIVA1
                End If
            
            End If
           
            'datos de retencion
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            If vParamAplic.ContabilidadNueva Then Sql = Sql & "0"
            
            ' Antigua: numdiari,fechaent,numasien,nodeducible)
            If Not vParamAplic.ContabilidadNueva Then Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
            
            Cad = Cad & "(" & Sql & ")"
            
            If vParamAplic.ContabilidadNueva Then
                Sql = "INSERT INTO factpro(numserie,numregis,fecfactu,anofactu,fecharec,fecliqpr,numfactu,codmacta,observa,nommacta,"
                Sql = Sql & "dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,codforpa,codopera,codconce340,codintra,"
                Sql = Sql & "totbases,totbasesret,totivas,totfacpr,retfacpr , trefacpr, cuereten, tiporeten)"
                Sql = Sql & " VALUES " & Cad
                ConnConta.Execute Sql
            Else
                'Insertar en la contabilidad
                Sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,fecliqpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                Sql = Sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                Sql = Sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,nodeducible) "
                Sql = Sql & " VALUES " & Cad
                ConnConta.Execute Sql
            End If
            
            If vParamAplic.ContabilidadNueva Then
                'Las  lineas de IVA
                Sql = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
                Sql = Sql & " VALUES " & CadenaInsertFaclin2
                ConnConta.Execute Sql
            End If
      
            
            'a�adido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            Sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            Sql = Sql & ",'" & DevNombreSQL(Rs!NumFactu) & " @ " & Format(Rs!FecFactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!nomprove) & "'," & Rs!codProve & ")"
            conn.Execute Sql
            
            
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

Public Function PasarFacturaTrans(cadwhere As String, CodCCost As String, FechaFin As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariagro.tcafpc --> conta.cabfactprov
' ariagro.tlifpc --> conta.linfactprov
'Actualizar la tabla ariges.scafpc.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim Sql As String
Dim Mc As Contadores


    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
        
    
    Set Mc = New Contadores
    
    '---- Insertar en la conta Cabecera Factura
    b = InsertarCabFactTrans(cadwhere, cadMen, Mc, FechaFin, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        CCoste = CodCCost
        '---- Insertar lineas de Factura en la Conta
        If Not vParamAplic.ContabilidadNueva Then
            b = InsertarLinFact_new("tcafpc", cadwhere, cadMen, Mc.Contador)
        Else
            b = InsertarLinFact_newContaNueva("tcafpc", cadwhere, cadMen, Mc.Contador)
        
        End If
        cadMen = "Insertando Lin. Factura: " & cadMen

        If b Then
            If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaProv(vContaFra.NumeroFactura, vContaFra.Anofac)
        
            '---- Poner intconta=1 en ariges.scafac
            b = ActualizarCabFact("tcafpc", cadwhere, cadMen)
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
            InsertarTMPErrFac cadMen, cadwhere
'            SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
'            SQL = SQL & " Select *," & DBSet(Mid(cadMen, 1, 200), "T") & " as error From tmpFactu "
'            SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
'            Conn.Execute SQL
        End If
    End If
End Function

Private Function InsertarCabFactTrans(cadwhere As String, cadErr As String, ByRef Mc As Contadores, FechaFin As String, ByRef vCF As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el a�o de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim Sql As String
Dim Sql5 As String
Dim Tipo As Byte
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Nulo4 As String
Dim TipoOpera As Integer
Dim Aux As String
Dim Sql2 As String
Dim ImporAux As Currency

Dim CadenaInsertFaclin2     As String

    On Error GoTo EInsertar
       
    
    Sql = " SELECT fecfactu,year(fecrecep) as anofacpr,fecrecep,numfactu,agencias.codmacta,"
    Sql = Sql & "tcafpc.dtoppago,tcafpc.dtognral,baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    Sql = Sql & "totalfac,tipoiva1,tipoiva2,tipoiva3, retfacpr, trefacpr, agencias.codtrans, agencias.nomtrans, "
    Sql = Sql & " tcafpc.domtrans,tcafpc.codpobla,tcafpc.pobtrans,tcafpc.protrans,tcafpc.niftrans,tcafpc.codforpa "
    Sql = Sql & " FROM " & "tcafpc "
    Sql = Sql & "INNER JOIN " & "agencias ON tcafpc.codtrans=agencias.codtrans "
    Sql = Sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
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
            '[Monica]09/06/2017: antes era baseiva2 y baseiva3
            If DBLet(Rs!TipoIVA2, "N") = "0" Then Nulo2 = "S"
            If DBLet(Rs!TipoIVA3, "N") = "0" Then Nulo3 = "S"
            If DBLet(Rs!trefacpr, "N") = "0" Then Nulo4 = "S"
            
            Sql = ""
            If vParamAplic.ContabilidadNueva Then Sql = "'" & SerieFraPro & "',"
            
            Sql = Sql & Mc.Contador & "," & DBSet(Rs!FecFactu, "F") & "," & Rs!anofacpr & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!NumFactu, "T") & "," & DBSet(Rs!Codmacta, "T") & "," & ValorNulo & ","
            
            If Not vParamAplic.ContabilidadNueva Then
                Sql = Sql & DBSet(Rs!BaseIVA1, "N") & "," & DBSet(Rs!BaseIVA2, "N", "S") & "," & DBSet(Rs!BaseIVA3, "N", "S") & ","
                Sql = Sql & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", Nulo2) & "," & DBSet(Rs!porciva3, "N", Nulo3) & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & DBSet(Rs!impoiva2, "N", Nulo2) & "," & DBSet(Rs!impoiva3, "N", Nulo3) & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!TipoIVA1, "N") & "," & DBSet(Rs!TipoIVA2, "N", Nulo2) & "," & DBSet(Rs!TipoIVA3, "N", Nulo3) & ",0,"
            Else
            
                Sql = Sql & DBSet(Rs!NomTrans, "T") & "," & DBSet(Rs!domtrans, "T", "S") & ","
                Sql = Sql & DBSet(Rs!codPobla, "T", "S") & "," & DBSet(Rs!pobtrans, "T", "S") & "," & DBSet(Rs!protrans, "T", "S") & ","
                Sql = Sql & DBSet(Rs!NIFTrans, "F", "S") & ",'ES',"
                Sql = Sql & DBSet(Rs!Codforpa, "N") & ","
                
                TipoOpera = 0
                
                Aux = "0"
                
                'codopera,codconce340,codintra
                Sql = Sql & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                
                
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
                Sql = Sql & DBSet(ImporAux, "N") & ","
                
                '[Monica]18/07/2017: en el caso de que tenga retencion grabamos la base de retencion
                If DBLet(Rs!trefacpr, "N") = "0" Then
                    Sql = Sql & ValorNulo & ","
                Else
                    Sql = Sql & DBSet(Rs!BaseIVA1, "N") & ","
                End If
                
                'totivas
                ImporAux = Rs!impoiva1 + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
                Sql = Sql & DBSet(ImporAux, "N") & "," & DBSet(Rs!TotalFac, "N") & ","
                        
                  
'                EsFacturaIntracom2 = ""
'                If DBLet(Rs!tipprove, "N") = 1 Then
'                    'OK es intracomunitaria
'                    EsFacturaIntracom2 = Rs!TipoIVA1
'                End If
                
            End If
            
            
            
            Sql = Sql & DBSet(Rs!retfacpr, "N", Nulo4) & "," & DBSet(Rs!trefacpr, "N", Nulo4) & ","
            
            If DBSet(Rs!retfacpr, "N", Nulo4) = ValorNulo And DBSet(Rs!trefacpr, "N", Nulo4) = ValorNulo Then
                Sql = Sql & ValorNulo & ","
            Else
                Sql5 = "select tipo from tcafpc where " & cadwhere
                Tipo = DevuelveValor(Sql5) ' 0=transportista 1=comisionista
            
                Select Case Tipo
                    Case 0 ' tranportista
                        Sql = Sql & DBSet(vParamAplic.CtaTraReten, "T") & ","
                    Case 1 ' comisionista
                        Sql = Sql & DBSet(vParamAplic.CtaComReten, "T") & ","
                End Select
            End If
            
            If vParamAplic.ContabilidadNueva Then
                If DBSet(Rs!trefacpr, "N", Nulo4) = ValorNulo Then
                    Sql = Sql & "0"
                Else
                    Sql = Sql & "1"
                End If
            End If
            
            If Not vParamAplic.ContabilidadNueva Then Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
            
            Cad = Cad & "(" & Sql & ")"
            
            If vParamAplic.ContabilidadNueva Then
                Sql = "INSERT INTO factpro(numserie,numregis,fecfactu,anofactu,fecharec,fecliqpr,numfactu,codmacta,observa,nommacta,"
                Sql = Sql & "dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,codforpa,codopera,codconce340,codintra,"
                Sql = Sql & "totbases,totbasesret,totivas,totfacpr,retfacpr , trefacpr, cuereten, tiporeten) values "
            
                ConnConta.Execute Sql & Cad
            
            Else
                'Insertar en la contabilidad
                Sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,fecliqpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                Sql = Sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                Sql = Sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,nodeducible) "
                Sql = Sql & " VALUES " & Cad
                ConnConta.Execute Sql
            End If
            
            If vParamAplic.ContabilidadNueva Then
                'Las  lineas de IVA
                Sql = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
                Sql = Sql & " VALUES " & CadenaInsertFaclin2
                ConnConta.Execute Sql
                
            End If
            
            
            'a�adido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            Sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            Sql = Sql & ",'" & DevNombreSQL(Rs!NumFactu) & " @ " & Format(Rs!FecFactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!NomTrans) & "'," & Rs!codTrans & ")"
            conn.Execute Sql
            
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
Public Function InsertarEnTesoreriaNewFac(cadwhere As String, CtaBan As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim b As Boolean
Dim Sql As String, text33csb As String, text41csb As String
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
    Sql = "select * from facturas where " & cadwhere
    Rsx.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
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
                Sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & DBSet(Rsx!Codforpa, "N")
                Set rsVenci = New ADODB.Recordset
                rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
                            Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                            Sql = Sql & " ctabanc1, iban, fecultco, impcobro, "
                            Sql = Sql & " text33csb, text41csb, agente, "
                            Sql = Sql & " nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais) "
                        Else
                            'Insertamos en la tabla scobro de la CONTA
                            Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                            Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                            Sql = Sql & " text33csb, text41csb, agente"
                            '[Monica]22/11/2013: Tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                Sql = Sql & ", iban) "
                            Else
                                Sql = Sql & ") "
                            End If
                        End If
                        Sql = Sql & " VALUES " & CadValues2
                        ConnConta.Execute Sql
                    
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
                Sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & DBSet(Rsx!Codforpa, "N")
                Set rsVenci = New ADODB.Recordset
                rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
                Sql7 = "select 0 tipo, numalbar, numlinealbar numlinea, impornet importe, codigiva from facturas_variedad where " & Replace(cadwhere, "facturas", "facturas_variedad")
                Sql7 = Sql7 & " union "
                '[Monica]11/02/2013: quieren en la referencia del iva el numero de albaran que pongan
                If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                    Sql7 = Sql7 & " select 1 tipo, numalbar numalbar, 0 numlinea, importel importe, codigiva from facturas_envases where " & Replace(cadwhere, "facturas", "facturas_envases")
                Else
                    Sql7 = Sql7 & " select 1 tipo,0, numlinea numlinea, importel importe, codigiva from facturas_envases where " & Replace(cadwhere, "facturas", "facturas_envases")
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
                        CadValues2 = CadValues2 & "'IVA VARIEDAD'," & DBSet(Format(DBLet(Rsx7!NumAlbar, "N"), "0000000"), "T") & "," & DBSet(Format(DBLet(Rsx7!NumLinea, "N"), "000"), "T")
                    Else
                        '[Monica]11/02/2013: metemos en la referencia el nro de albaran que hayan metido
                        If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                            CadValues2 = CadValues2 & "'IVA ENVASE'," & DBSet(DBLet(Rsx7!NumAlbar, "N"), "T") & "," & ValorNulo
                        Else
                            CadValues2 = CadValues2 & "'IVA ENVASE'," & DBSet(Format(DBLet(Rsx7!NumLinea, "N"), "000"), "T") & "," & ValorNulo
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
                    Sql7 = "select sum(importel) from facturas_envases where " & Replace(cadwhere, "facturas", "facturas_envases")
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
                    Sql7 = "select numlinea, numalbar, importel from facturas_envases where " & Replace(cadwhere, "facturas", "facturas_envases")
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
                                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & text33csb & "," & DBSet(text41csb, "T") & ",1,'ENVASES'," & DBSet(Rsx7.Fields(1).Value, "T") & "," & ValorNulo & ","  '& "),"
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
                Sql7 = "select * from facturas_variedad where " & Replace(cadwhere, "facturas", "facturas_variedad")
                Set Rsx7 = New ADODB.Recordset
                
                Rsx7.Open Sql7, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not Rsx7.EOF
                    '[Monica]09/11/2017: para el caso de que el importe sea 0 tambien se introduce, pq si no cuando se recibe la factura de
                    '                    no lo encuentra.
                    '                    quito la condicion de que el importe sea <> 0 PARA EL CASO DE PICASSENT S�LO !!!!!!!!!
                    If DBLet(Rsx7!impornet, "N") <> 0 Or (DBLet(Rsx7!impornet, "N") = 0 And vParamAplic.Cooperativa = 2) Then
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
                        Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        Sql = Sql & "ctabanc1, iban, fecultco, impcobro, "
                        Sql = Sql & " text33csb, text41csb, agente, referencia, referencia1, referencia2," ') "
                        Sql = Sql & "nomclien,domclien,pobclien,cpclien,proclien,nifclien,codpais"
                        Sql = Sql & ") "
                    Else
                        'Insertamos en la tabla scobro de la CONTA
                        Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                        Sql = Sql & " text33csb, text41csb, agente, referencia, referencia1, referencia2" ') "
                        
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            Sql = Sql & ", iban) "
                        Else
                            Sql = Sql & ") "
                        End If
                    End If
                    Sql = Sql & " VALUES " & CadValues2
                    ConnConta.Execute Sql
                    
                    
                    If ImporteACompensar <> 0 Then
                        If vParamAplic.ContabilidadNueva Then
                            Sql = "update cobros set impvenci = impvenci + " & DBSet(ImporteACompensar, "N")
                            Sql = Sql & " where numserie = " & DBSet(letraser, "T")
                            Sql = Sql & " and numfactu = " & DBSet(Rsx!NumFactu, "N")
                            Sql = Sql & " and fecfactu = " & DBSet(Rsx!FecFactu, "F")
                            Sql = Sql & " and numorden = " & DBSet(LineaAModificar, "N")
                            
                            ConnConta.Execute Sql
                        
                        Else
                            Sql = "update scobro set impvenci = impvenci + " & DBSet(ImporteACompensar, "N")
                            Sql = Sql & " where numserie = " & DBSet(letraser, "T")
                            Sql = Sql & " and codfaccl = " & DBSet(Rsx!NumFactu, "N")
                            Sql = Sql & " and fecfaccl = " & DBSet(Rsx!FecFactu, "F")
                            Sql = Sql & " and numorden = " & DBSet(LineaAModificar, "N")
                            
                            ConnConta.Execute Sql
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
Public Function InsertarEnTesoreriaNewFacSoc(cadwhere As String, CtaBan As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim b As Boolean
Dim Sql As String, text33csb As String, text41csb As String
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
    Sql = "select * from facturassocio where " & cadwhere
    Rsx.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
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
            Sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & DBSet(Rsx!Codforpa, "N")
            Set rsVenci = New ADODB.Recordset
            rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

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
                        CadValues2 = CadValues2 & DBSet(Rs4!nomsocio, "T") & "," & DBSet(Rs4!dirsocio, "T") & "," & DBSet(Rs4!pobsocio, "T") & "," & DBSet(Rs4!codPostal, "T") & "," & DBSet(Rs4!prosocio, "T") & "," & DBSet(Rs4!nifsocio, "T") & ",'ES'"
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
                            CadValues2 = CadValues2 & DBSet(Rs4!nomsocio, "T") & "," & DBSet(Rs4!dirsocio, "T") & "," & DBSet(Rs4!pobsocio, "T") & "," & DBSet(Rs4!codPostal, "T") & "," & DBSet(Rs4!prosocio, "T") & "," & DBSet(Rs4!nifsocio, "T") & ",'ES'"
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
                        Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        Sql = Sql & "ctabanc1, iban, fecultco, impcobro, "
                        Sql = Sql & " text33csb, text41csb, agente,nomclien,domclien,pobclien,cpclien,proclien,nifclien,codpais) " ') "
                    
                    Else
                        Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                        Sql = Sql & " text33csb, text41csb, agente" ') "
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            Sql = Sql & ", iban) "
                        Else
                            Sql = Sql & ") "
                        End If
                    End If
                    
                    Sql = Sql & " VALUES " & CadValues2
                    ConnConta.Execute Sql
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
Dim Sql As String

    Sql = "select count(*) from tmpfactu where codtipom = 'EAC'"
    
    HayFacturasACuenta = (TotalRegistros(Sql) <> 0)

End Function


Public Function InsertarAsientoDiario(FecEnt As String, NDiario As String, CtaContra As String, NLiq As String, FecLiq As String, cadErr As String)
Dim Sql As String
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
Dim Cad As String
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
        
            Sql = "select distinct * from tmpinformes, ariconta" & vParamAplic.NumeroConta & ".cobros cc "
            Sql = Sql & " where tmpinformes.codusu = " & vUsu.Codigo
            Sql = Sql & " and tmpinformes.nombre1 = cc.numserie "
            Sql = Sql & " and tmpinformes.importe1 = cc.numfactu "
            Sql = Sql & " and tmpinformes.fecha1 = cc.fecfactu "
            Sql = Sql & " and tmpinformes.importe2 = cc.numorden "
            Sql = Sql & " order by importe1, fecha1, importe2 "
        
        Else
    
            Sql = "select distinct * from tmpinformes, conta" & vParamAplic.NumeroConta & ".scobro cc "
            Sql = Sql & " where codusu = " & vUsu.Codigo
            Sql = Sql & " and tmpinformes.nombre1 = cc.numserie "
            Sql = Sql & " and tmpinformes.importe1 = cc.codfaccl "
            Sql = Sql & " and tmpinformes.fecha1 = cc.fecfaccl "
            Sql = Sql & " and tmpinformes.importe2 = cc.numorden "
            Sql = Sql & " order by importe1, fecha1, importe2 "
            
        End If
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        i = 0
        ImporteD = 0
        ImporteH = 0
    
        b = True
        
        While Not Rs.EOF And b
            If vParamAplic.ContabilidadNueva Then
                numdocum = Rs!numSerie & Format(Rs!NumFactu, "0000000")
                tipoF = DevuelveValor("select tipforpa from formapago where codforpa = " & DBSet(Rs!Codforpa, "N"))
                
                Conceptoh = "conhacli"
                Conceptod = DevuelveDesdeBDNew(cConta, "tipofpago", "condecli", "tipoformapago", tipoF, "N", Conceptoh)
            Else
                numdocum = Rs!numSerie & Format(Rs!Codfaccl, "0000000")
                tipoF = DevuelveValor("select tipoforp from forpago where codforpa = " & DBSet(Rs!Codforpa, "N"))
            
                Conceptoh = "conhacli"
                Conceptod = DevuelveDesdeBDNew(cConta, "stipoformapago", "condecli", "tipoformapago", tipoF, "N", Conceptoh)
            End If
            
            Amplia = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", CStr(Conceptod), "N"))
            ampliaciond = Amplia & " " & Format(Rs!referencia1, "0000000") & "-" & DBLet(Rs!referencia2)
            
            Amplia = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", CStr(Conceptoh), "N"))
            ampliacionh = Amplia & " " & Format(Rs!referencia1, "0000000") & "-" & DBLet(Rs!referencia2)
                
            i = i + 1
            
            Cad = DBSet(NDiario, "N") & "," & DBSet(FecEnt, "F") & "," & DBSet(Mc.Contador, "N") & ","
            Cad = Cad & DBSet(i, "N") & "," & DBSet(Rs!Codmacta, "T") & "," & DBSet(numdocum, "T") & ","
            
            ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
            If DBLet(Rs!ImpVenci, "N") > 0 Then
                ' importe al haber en positivo
                Cad = Cad & DBSet(Conceptoh, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                Cad = Cad & DBSet(Rs!ImpVenci, "N") & "," & ValorNulo & "," & DBSet(CtaContra, "T") & "," & ValorNulo & ",0"
            
                ImporteH = ImporteH + CCur(DBLet(Rs!ImpVenci, "N"))
            Else
                ' importe al debe en positivo cambiamos signo
                Cad = Cad & DBSet(Conceptod, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(DBLet(Rs!ImpVenci, "N") * (-1), "N") & ","
                Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaContra, "T") & "," & ValorNulo & ",0"
            
                ImporteD = ImporteD + CCur(DBLet(Rs!ImpVenci, "N") * (-1))
            End If
            
            ' como pasan a estar cobrados se tienen que grabar la serie, factura y fechafactu
            If vParamAplic.ContabilidadNueva Then
                Cad = Cad & "," & DBSet(Rs!numSerie, "T") & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!numorden, "N") & "," & DBSet(tipoF, "N")
            End If
            
            Cad = "(" & Cad & ")"
            
            b = InsertarLinAsientoDia(Cad, cadMen)
            cadMen = "Insertando Lin. Asiento: " & i
            
            
            Rs.MoveNext
            
        Wend
    
        If b Then
    
            i = i + 1
                    
            numdocum = Format(NLiq, "0000000")
                    
            ' el Total es sobre la cuenta del cliente
            Cad = DBSet(NDiario, "N") & "," & DBSet(FecEnt, "F") & "," & DBSet(Mc.Contador, "N") & ","
            Cad = Cad & DBSet(i, "N") & ","
            Cad = Cad & DBSet(CtaContra, "T") & "," & DBSet(numdocum, "T") & ","
                
            Total = ImporteH - ImporteD
                
            ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
            If Total > 0 Then
                ' importe al debe en positivo
                Cad = Cad & DBSet(Conceptod, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Total, "N") & ","
                Cad = Cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
            Else
                ' importe al haber en positivo, cambiamos el signo
                Cad = Cad & DBSet(Conceptoh, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                Cad = Cad & DBSet(Total * (-1), "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
            End If
            
            
            If vParamAplic.ContabilidadNueva Then
                Cad = Cad & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
            End If
            
            Cad = "(" & Cad & ")"
            
            b = InsertarLinAsientoDia(Cad, cadMen)
            cadMen = "Insertando Lin. Asiento: " & i
        End If
        
        If b Then b = EliminarCobros(cadMen, FecEnt)
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String

    On Error GoTo EInsertar
       
    
    If vParamAplic.ContabilidadNueva Then
        Cad = Format(Diario, "00") & ", " & DBSet(fecha, "F") & "," & Format(Asiento, "000000") & ","
        Cad = Cad & DBSet(Obs, "T") & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARIAGRO COMERCIAL'"
        
        Cad = "(" & Cad & ")"
    
        'Insertar en la contabilidad
        Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) "
        Sql = Sql & " VALUES " & Cad
        ConnConta.Execute Sql
    Else
        Cad = Format(Diario, "00") & ", " & DBSet(fecha, "F") & "," & Format(Asiento, "000000") & ","
        Cad = Cad & "0," & ValorNulo & "," & DBSet(Obs, "T")
        Cad = "(" & Cad & ")"
    
        'Insertar en la contabilidad
        Sql = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) "
        Sql = Sql & " VALUES " & Cad
        ConnConta.Execute Sql
    End If
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabAsientoDia = False
        cadErr = Err.Description
    Else
        InsertarCabAsientoDia = True
    End If
End Function



Public Function InsertarLinAsientoDia(Cad As String, cadErr As String) As Boolean
' el Tipo me indica desde donde viene la llamada
' tipo = 0 srecau.codmacta
' tipo = 1 scaalb.codmacta

Dim Rs As ADODB.Recordset
Dim Aux As String
Dim Sql As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    If vParamAplic.ContabilidadNueva Then
        Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
        Sql = Sql & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada, numserie, numfaccl, fecfactu, numorden, tipforpa) "
        Sql = Sql & " VALUES " & Cad
    
    Else
 
        Sql = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
        Sql = Sql & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
        Sql = Sql & " VALUES " & Cad
        
    End If
    ConnConta.Execute Sql

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinAsientoDia = False
        cadErr = Err.Description
    Else
        InsertarLinAsientoDia = True
    End If
End Function



Private Function EliminarCobros(cadErr As String, FechaCobro As String) As Boolean

Dim Rs As ADODB.Recordset
Dim Aux As String
Dim Sql As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    If vParamAplic.ContabilidadNueva Then
    
'        sql = "DELETE FROM ariconta" & vParamAplic.NumeroConta & ".cobros where (numserie,numfactu,fecfactu,numorden) in "
'        sql = sql & " (select nombre1, importe1, fecha1, importe2 from tmpinformes where codusu = " & vUsu.Codigo & ")"
        
        Sql = "update ariconta" & vParamAplic.NumeroConta & ".cobros set "
        Sql = Sql & " impcobro = impvenci + coalesce(gastos,0) "
        Sql = Sql & ",fecultco = " & DBSet(FechaCobro, "F")
        Sql = Sql & ",situacion = 1"
        Sql = Sql & " where (numserie,numfactu,fecfactu,numorden) in "
        Sql = Sql & " (select nombre1, importe1, fecha1, importe2 from tmpinformes where codusu = " & vUsu.Codigo & ")"
    
    Else
 
        Sql = "DELETE FROM conta" & vParamAplic.NumeroConta & ".scobro where (numserie,codfaccl,fecfaccl,numorden) in "
        Sql = Sql & " (select nombre1, importe1, fecha1, importe2 from tmpinformes where codusu = " & vUsu.Codigo & ")"
    
    End If
    
    conn.Execute Sql

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
Dim Sql As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

 
    Sql = "UPDATE anecoop_pago SET idcontab = 1 where (expediente_id, expediente_pagoid) in "
    Sql = Sql & " (select nombre1, nombre2 from tmpinformes2 where codusu = " & vUsu.Codigo & ")"
    
    conn.Execute Sql

EInLinea:
    If Err.Number <> 0 Then
        MarcarRegistros = False
        cadErr = Err.Description
    Else
        MarcarRegistros = True
    End If
End Function



'----------------------------------------------------------------------
' FACTURAS DE COMPRA
'----------------------------------------------------------------------

Public Function PasarFacturaCompras(cadwhere As String, CodCCost As String, FechaFin As String, ByRef vContaFra As cContabilizarFacturas, CtaPrev As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariges.scafpc --> conta.cabfactprov
' ariges.slifpc --> conta.linfactprov
'Actualizar la tabla ariges.scafpc.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim Sql As String
Dim Mc As Contadores
Dim FraIntraCom2 As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
        
    
    Set Mc = New Contadores
    
    '---- Insertar en la conta Cabecera Factura
    b = InsertarCabFactCompra(cadwhere, cadMen, Mc, FechaFin, vContaFra, FraIntraCom2)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
    
        b = InsertarEnTesoreriaCompras(cadwhere, CtaPrev, cadMen)
        cadMen = "Insertando en Tesoreria: " & cadMen
    
    
        CCoste = CodCCost
        '---- Insertar lineas de Factura en la Conta
        If Not vParamAplic.ContabilidadNueva Then
            b = InsertarLinFact_new("facturascom", cadwhere, cadMen, Mc.Contador)
        Else
            b = InsertarLinFact_newContaNueva("facturascom", cadwhere, cadMen, Mc.Contador, FraIntraCom2)
        End If
        cadMen = "Insertando Lin. Factura: " & cadMen

        If b Then
            If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaProv(vContaFra.NumeroFactura, vContaFra.Anofac)
            
            '---- Poner intconta=1 en ariges.scafac
            b = ActualizarCabFact("facturascom", cadwhere, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaCompras = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaCompras = False
        If Not b Then
            InsertarTMPErrFac cadMen, cadwhere
'            SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
'            SQL = SQL & " Select *," & DBSet(Mid(cadMen, 1, 200), "T") & " as error From tmpFactu "
'            SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
'            Conn.Execute SQL
        End If
    End If
End Function


Private Function InsertarCabFactCompra(cadwhere As String, cadErr As String, ByRef Mc As Contadores, FechaFin As String, ByRef vCF As cContabilizarFacturas, ByRef EsFacturaIntracom2 As String) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el a�o de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Intracom As Integer

Dim TipoOpera As Byte
Dim CadenaInsertFaclin2     As String
Dim ImporAux As Currency

Dim Aux As String
Dim Sql2 As String

    On Error GoTo EInsertar
       
    
    Sql = " SELECT fecfactu,year(fecrecep) as anofacpr,fecrecep,numfactu,proveedor.codmacta,"
    Sql = Sql & "facturascom.dtoppago,facturascom.dtognral,baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    Sql = Sql & "totalfac,codiiva1,codiiva2,codiiva3,proveedor.codprove, proveedor.nomprove, proveedor.tipprove, facturascom.observac "
    Sql = Sql & ",proveedor.domprove,proveedor.codpobla,proveedor.pobprove,proveedor.proprove,proveedor.nifprove,facturascom.codforpa, proveedor.tipprove "
    Sql = Sql & " FROM " & "facturascom "
    Sql = Sql & "INNER JOIN " & "proveedor ON facturascom.codprove=proveedor.codprove "
    Sql = Sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
    
        If Mc.ConseguirContador("1", (Rs!FecRecep <= CDate(FechaFin) - 365), True) = 0 Then
        
            vCF.NumeroFactura = Mc.Contador
            vCF.Anofac = Year(DBLet(Rs!FecRecep))
            vvFecRecep = DBLet(Rs!FecRecep)
        
            '[Monica]27/07/2017
            vCF.observa = DBLet(Rs!Observac)
        
            
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            DtoPPago = Rs!DtoPPago
            DtoGnral = Rs!DtoGnral
            BaseImp = Rs!baseimp1 + CCur(DBLet(Rs!baseimp2, "N")) + CCur(DBLet(Rs!baseimp3, "N"))
            TotalFac = Rs!TotalFac
            AnyoFacPr = Rs!anofacpr
            
            Intracom = DBLet(Rs!tipprove, "N")
            If Intracom = 2 Then Intracom = 0
            
            Nulo2 = "N"
            Nulo3 = "N"
            '[Monica]09/06/2017: antes era baseiva2 y baseiva3
            If DBLet(Rs!codiiva2, "N") = "0" Then Nulo2 = "S"
            If DBLet(Rs!codiiva3, "N") = "0" Then Nulo3 = "S"
            Sql = ""
            If vParamAplic.ContabilidadNueva Then Sql = "'" & SerieFraPro & "',"
            
            Sql = Sql & Mc.Contador & "," & DBSet(Rs!FecFactu, "F") & "," & Rs!anofacpr & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!NumFactu, "T") & "," & DBSet(Rs!Codmacta, "T") & "," & DBSet(Rs!Observac, "T") & ","
            
            If Not vParamAplic.ContabilidadNueva Then
                Sql = Sql & DBSet(Rs!baseimp1, "N") & "," & DBSet(Rs!baseimp2, "N", "S") & "," & DBSet(Rs!baseimp3, "N", "S") & ","
                Sql = Sql & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", Nulo2) & "," & DBSet(Rs!porciva3, "N", Nulo3) & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & DBSet(Rs!impoiva2, "N", Nulo2) & "," & DBSet(Rs!impoiva3, "N", Nulo3) & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!codiiva1, "N") & "," & DBSet(Rs!codiiva2, "N", Nulo2) & "," & DBSet(Rs!codiiva3, "N", Nulo3) & "," & DBSet(Intracom, "N") & ","
            Else
                Sql = Sql & DBSet(Rs!nomprove, "T") & "," & DBSet(Rs!domprove, "T", "S") & ","
                Sql = Sql & DBSet(Rs!codPobla, "T", "S") & "," & DBSet(Rs!pobprove, "T", "S") & "," & DBSet(Rs!proprove, "T", "S") & ","
                Sql = Sql & DBSet(Rs!nifProve, "F", "S") & ",'ES',"
                Sql = Sql & DBSet(Rs!Codforpa, "N") & ","
                
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
'[Monica]08/06/2017: no s� cuando es rectificativa en proveedores
'                    If Rs!TotalFac < 0 Then
'                        Aux = "D"
'                    Else
                        If Not IsNull(Rs!codiiva2) Then Aux = "C"
'                    End If
                
                Case 1
                    Aux = "P"
                
                Case 4
                    Aux = "I"
                End Select
                
                'codopera,codconce340,codintra
                Sql = Sql & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                
                
                'para las lineas
                'factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
                'IVA 1, siempre existe
                Aux = "'" & SerieFraPro & "'," & Mc.Contador & "," & DBSet(Rs!FecRecep, "F") & "," & Rs!anofacpr & ","
                
                Sql2 = Aux & "1," & DBSet(Rs!baseimp1, "N") & "," & Rs!codiiva1 & "," & DBSet(Rs!porciva1, "N") & ","
                Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
                vTipoIva(0) = Rs!codiiva1
                vPorcIva(0) = Rs!porciva1
                vPorcRec(0) = 0
                vImpIva(0) = DBLet(Rs!impoiva1, "N")
                vImpRec(0) = 0
                vBaseIva(0) = DBLet(Rs!baseimp1, "N")
                
                vTipoIva(1) = 0: vTipoIva(2) = 0
                
                If Not IsNull(Rs!porciva2) Then
                    Sql2 = Aux & "2," & DBSet(Rs!baseimp2, "N") & "," & Rs!codiiva2 & "," & DBSet(Rs!porciva2, "N") & ","
                    Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva2, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                    vTipoIva(1) = Rs!codiiva2
                    vPorcIva(1) = Rs!porciva2
                    vPorcRec(1) = 0
                    vImpIva(1) = DBLet(Rs!impoiva2, "N")
                    vImpRec(1) = 0
                    vBaseIva(1) = DBLet(Rs!baseimp2, "N")
                
                End If
                If Not IsNull(Rs!porciva3) Then
                    Sql2 = Aux & "3," & DBSet(Rs!baseimp3, "N") & "," & Rs!codiiva3 & "," & DBSet(Rs!porciva3, "N") & ","
                    Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva3, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                    vTipoIva(2) = Rs!codiiva3
                    vPorcIva(2) = Rs!porciva3
                    vPorcRec(2) = 0
                    vImpIva(2) = DBLet(Rs!impoiva3, "N")
                    vImpRec(2) = 0
                    vBaseIva(2) = DBLet(Rs!baseimp3, "N")
                End If
                
                    
                    
                'Los totales
                'totbases,totbasesret,totivas,totrecargo,totfacpr,
                ImporAux = Rs!baseimp1 + DBLet(Rs!baseimp2, "N") + DBLet(Rs!baseimp3, "N")
                Sql = Sql & DBSet(ImporAux, "N") & "," & ValorNulo & ","
                'totivas
                ImporAux = Rs!impoiva1 + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
                Sql = Sql & DBSet(ImporAux, "N") & "," & DBSet(Rs!TotalFac, "N") & ","
                        
                  
                EsFacturaIntracom2 = ""
                If DBLet(Rs!tipprove, "N") = 1 Then
                    'OK es intracomunitaria
                    EsFacturaIntracom2 = Rs!codiiva1
                End If
            
            End If
           
            'datos de retencion
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            If vParamAplic.ContabilidadNueva Then Sql = Sql & "0"
            
            ' Antigua: numdiari,fechaent,numasien,nodeducible)
            If Not vParamAplic.ContabilidadNueva Then Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
            
            Cad = Cad & "(" & Sql & ")"
            
            If vParamAplic.ContabilidadNueva Then
                Sql = "INSERT INTO factpro(numserie,numregis,fecfactu,anofactu,fecharec,fecliqpr,numfactu,codmacta,observa,nommacta,"
                Sql = Sql & "dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,codforpa,codopera,codconce340,codintra,"
                Sql = Sql & "totbases,totbasesret,totivas,totfacpr,retfacpr , trefacpr, cuereten, tiporeten)"
                Sql = Sql & " VALUES " & Cad
                ConnConta.Execute Sql
            Else
                'Insertar en la contabilidad
                Sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,fecliqpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                Sql = Sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                Sql = Sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,nodeducible) "
                Sql = Sql & " VALUES " & Cad
                ConnConta.Execute Sql
            End If
            
            If vParamAplic.ContabilidadNueva Then
                'Las  lineas de IVA
                Sql = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
                Sql = Sql & " VALUES " & CadenaInsertFaclin2
                ConnConta.Execute Sql
            End If
      
            
            'a�adido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            Sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            Sql = Sql & ",'" & DevNombreSQL(Rs!NumFactu) & " @ " & Format(Rs!FecFactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!nomprove) & "'," & Rs!codProve & ")"
            conn.Execute Sql
            
            
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactCompra = False
        cadErr = Err.Description
    Else
        InsertarCabFactCompra = True
    End If
End Function

Private Function InsertarEnTesoreriaCompras(cadwhere As String, CuentaPrev As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim Sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single
Dim i As Byte
Dim vvIban As String
Dim vProv As CProveedor
Dim Sql2 As String
Dim Rs2 As ADODB.Recordset
Dim ForPago As String

    On Error GoTo EInsertarTesoreria

    InsertarEnTesoreriaCompras = False
    CadValues2 = ""

    Sql2 = "select * from facturascom where " & cadwhere
    Set Rs2 = New ADODB.Recordset
    Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs2.EOF Then

        ForPago = DBLet(Rs2!Codforpa)
        
        Set vProv = New CProveedor
        If Not vProv.LeerDatos(Rs2!codProve) Then
            Set vProv = Nothing
            Exit Function
        End If
        
        'Obtener el N� de Vencimientos de la forma de pago
        Sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & ForPago
        Set rsVenci = New ADODB.Recordset
        rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not rsVenci.EOF Then
            If rsVenci!numerove > 0 Then
                'Obtener los dias de pago de la tabla de parametros: spara1
                Sql = " SELECT  diapago1, diapago2, diapago3,mesnogir "
                Sql = Sql & " FROM sparam "
                Sql = Sql & " WHERE codparam=1"
                Set Rs = New ADODB.Recordset
                Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Not Rs.EOF Then
                   'vamos creando la cadena para insertar en spagosp de la CONTA
                   CadValuesAux2 = "("
                   If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & DBSet(SerieFraPro, "T") & ","
                   CadValuesAux2 = CadValuesAux2 & "'" & vProv.CuentaCble & "', " & DBSet(Rs2!NumFactu, "T") & ", '" & Format(Rs2!FecFactu, FormatoFecha) & "', "
                  
                  'Primer Vencimiento
                  '------------------------------------------------------------
                  i = 1
                  'FECHA VTO
                  FecVenci = CDate(Rs2!FecFactu)
                  '=== Modificado: Laura 23/01/2007
    '              FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                  FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                  '==================================
                  'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                  FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
    
                  'Comprobar si  tiene mes a no girar
                  FecVenci1 = FecVenci
                  If DBSet(Rs!mesnogir, "N") <> 0 Then
                      FecVenci1 = ComprobarMesNoGira(FecVenci1, DBSet(Rs!mesnogir, "N"), DBSet(0, "N"), Rs!DiaPago1, Rs!DiaPago2, Rs!DiaPago3)
                  End If
                 
                  CadValues2 = CadValuesAux2 & i
                  CadValues2 = CadValues2 & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                  'IMPORTE del Vencimiento
                  If rsVenci!numerove = 1 Then
                        ImpVenci = DBLet(Rs2!TotalFac, "N")
                  Else
                        ImpVenci = Round(DBLet(Rs2!TotalFac, "N") / rsVenci!numerove, 2)
                        'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                        If ImpVenci * rsVenci!numerove <> DBLet(Rs2!TotalFac) Then
                            ImpVenci = Round(ImpVenci + (DBLet(Rs2!TotalFac) - ImpVenci * rsVenci!numerove), 2)
                        End If
                  End If
                  CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CuentaPrev, "T") & ","
                  If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(vProv.Iban, "") & MiFormat(vProv.Banco, "0000") & MiFormat(vProv.Sucursal, "0000") & MiFormat(vProv.DigControl, "00") & MiFormat(vProv.CuentaBan, "0000000000")
                        
                        CadValues2 = CadValues2 & DBSet(vvIban, "T", "S") & ","
                  Else
                    'David. Para que ponga la cuenta bancaria (SI LA tiene)
                    CadValues2 = CadValues2 & DBSet(vProv.Banco, "T", "S") & "," & DBSet(vProv.Sucursal, "T", "S") & ","
                    CadValues2 = CadValues2 & DBSet(vProv.DigControl, "T", "S") & "," & DBSet(vProv.CuentaBan, "T", "S") & ","
                    '[Monica]22/11/2013: Team iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                         CadValues2 = CadValues2 & DBSet(vProv.Iban, "T", "S") & ","
                    End If
                  
                  End If
    
    
                  'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
                  Sql = "Factura num.: " & Rs2!NumFactu & "-" & Format(Rs2!FecFactu, "dd/mm/yyyy")
                  CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
                  Sql = "Vto a fecha: " & Format(FecVenci1, "dd/mm/yyyy")
                  CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'"
    
                  If vParamAplic.ContabilidadNueva Then
                      CadValues2 = CadValues2 & "," & DBSet(vProv.Nombre, "T") & "," & DBSet(vProv.Domicilio, "T") & "," & DBSet(vProv.Poblacion, "T") & "," & DBSet(vProv.CPostal, "T") & "," & DBSet(vProv.Provincia, "T") & "," & DBSet(vProv.NIF, "T") & ",'ES')"
                  Else
                      CadValues2 = CadValues2 & ")"
                  End If
    
    
                  'Resto Vencimientos
                  '--------------------------------------------------------------------
                  For i = 2 To rsVenci!numerove
                     'FECHA Resto Vencimientos
                      '==== Modificado: Laura 23/01/2007
                      'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                      FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                      '==================================================
                      'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                      FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
    
                      'Comprobar si tiene mes a no girar
                      FecVenci1 = FecVenci
                      If DBSet(Rs!mesnogir, "N") <> 0 Then
                            FecVenci1 = ComprobarMesNoGira(FecVenci1, DBSet(Rs!mesnogir, "N"), DBSet(0, "N"), Rs!DiaPago1, Rs!DiaPago2, Rs!DiaPago3)
                      End If
    
                      CadValues2 = CadValues2 & ", " & CadValuesAux2 & i & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
    
                      'IMPORTE Resto de Vendimientos
                      ImpVenci = Round(Rs2!TotalFac / rsVenci!numerove, 2)
    
                      CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", '" & CuentaPrev & "',"
                      
                      
                      'David. Para que ponga la cuenta bancaria (SI LA tiene)
                      If vParamAplic.ContabilidadNueva Then
                            vvIban = MiFormat(vProv.Iban, "") & MiFormat(vProv.Banco, "0000") & MiFormat(vProv.Sucursal, "0000") & MiFormat(vProv.DigControl, "00") & MiFormat(vProv.CuentaBan, "0000000000")
                            
                            CadValues2 = CadValues2 & DBSet(vvIban, "T", "S") & ","
                      Else
                            CadValues2 = CadValues2 & DBSet(vProv.Banco, "T", "S") & "," & DBSet(vProv.Sucursal, "T", "S") & ","
                            CadValues2 = CadValues2 & DBSet(vProv.DigControl, "T", "S") & "," & DBSet(vProv.CuentaBan, "T", "S") & ","
                      End If
                      '[Monica]22/11/2013: Team iban
                      If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & DBSet(vProv.Iban, "T", "S") & ","
                      End If
    
                      Sql = "Factura num.: " & Rs2!NumFactu & "-" & Format(Rs2!FecFactu, "dd/mm/yyyy")
                      CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
                      Sql = "Vto a fecha: " & Format(FecVenci1, "dd/mm/yyyy")
                      CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'"
                      
                      If vParamAplic.ContabilidadNueva Then
                            CadValues2 = CadValues2 & "," & DBSet(vProv.Nombre, "T") & "," & DBSet(vProv.Domicilio, "T") & "," & DBSet(vProv.Poblacion, "T") & "," & DBSet(vProv.CPostal, "T") & "," & DBSet(vProv.Provincia, "T") & "," & DBSet(vProv.NIF, "T") & ",'ES')"
                      Else
                            CadValues2 = CadValues2 & ")"
                      End If
                  Next i
                End If
            End If
            Rs.Close
            Set Rs = Nothing
        End If
        rsVenci.Close
        Set rsVenci = Nothing
        
        'Grabar tabla spagop de la CONTABILIDAD
        '-------------------------------------------------
        If CadValues2 <> "" Then
            'antes de grabar en la spago comprobar que existe en conta.sforpa la
            'forma de pago de la factura. Sino existe insertarla
    
            'vemos si existe en la conta
            If vParamAplic.ContabilidadNueva Then
                CadValuesAux2 = DevuelveDesdeBDNew(cConta, "formapago", "codforpa", "codforpa", ForPago, "N")
            Else
                CadValuesAux2 = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", ForPago, "N")
            End If
            'si no existe la forma de pago en conta, insertamos la de ariges
            If CadValuesAux2 = "" Then
        '++
        
                Dim Sql8 As String
                Dim RS8 As ADODB.Recordset
        
                Sql8 = "select * from forpago where codforpa = " & DBSet(ForPago, "N")
                Set RS8 = New ADODB.Recordset
                RS8.Open Sql8, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RS8.EOF Then
                    'insertamos e sforpa de la CONTA
                    If vParamAplic.ContabilidadNueva Then
                        Sql8 = "INSERT INTO formapago(codforpa,nomforpa,tipforpa,numerove,primerve,restoven)"
                    Else
                        Sql8 = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
                    End If
                    Sql8 = Sql8 & " VALUES(" & DBSet(ForPago, "N") & ", " & DBSet(RS8!nomforpa, "T") & ", " & DBSet(RS8!tipoforp, "N")
                    If vParamAplic.ContabilidadNueva Then
                        Sql8 = Sql8 & "," & DBSet(RS8!numerove, "N") & "," & DBSet(RS8!primerve, "N") & "," & DBSet(RS8!restoven, "N") & ")"
                    Else
                        Sql8 = Sql8 & ")"
                    End If
                    ConnConta.Execute Sql8
                End If
                RS8.Close
                Set RS8 = Nothing
         '++
    
    
    
    
            End If
    
            'Insertamos en la tabla spagop de la CONTA
            'SQL = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1) "
            'David. Cuenta bancaria y descripcion textos
            If vParamAplic.ContabilidadNueva Then
                Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,iban,"
                Sql = Sql & "text1csb,text2csb,nomprove,domprove,pobprove,cpprove,proprove,nifprove,codpais) "
            
            Else
                Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba," 'text1csb,text2csb) "
                '[Monica]22/11/2013:Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    Sql = Sql & "iban,text1csb,text2csb) "
                Else
                    Sql = Sql & "text1csb,text2csb) "
                End If
            End If
        
            Sql = Sql & " VALUES " & CadValues2
            
            ConnConta.Execute Sql
        
        
        End If
    
        b = True
    End If
    
    Set vProv = Nothing
    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        b = False
        MenError = "Error al insertar en Tesoreria Compras: " & Err.Description
    End If
    InsertarEnTesoreriaCompras = b
End Function



