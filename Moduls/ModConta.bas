Attribute VB_Name = "ModConta"
Option Explicit

'=============================================================================
'   MODULO PARA ACCEDER A LOS DATOS DE LA CONTABILIDAD
'=============================================================================


'=============================================================================
'==========     CUENTAS
'=============================================================================
'LAURA
Public Function PonerNombreCuenta(ByRef Txt As TextBox, Modo As Byte, Optional clien As String) As String
'Obtener el nombre de una cuenta
Dim DevfrmCCtas As String
Dim cad As String

' ### [Monica] 07/09/2006 añadida la linea siguiente condicion vParamAplic.NumeroConta = 0
' para que no saque nada si no hay contabilidad
    If Not vParamAplic Is Nothing Then
        If vParamAplic.NumeroConta = 0 Then
            PonerNombreCuenta = ""
            Exit Function
        End If
    End If
    If Txt.Text = "" Then
         PonerNombreCuenta = ""
         Exit Function
    End If
    DevfrmCCtas = Txt.Text
    If CuentaCorrectaUltimoNivel(DevfrmCCtas, cad) Then
        ' ### [Monica] 07/09/2006
        If InStr(cad, "No existe la cuenta") > 0 Then
            Txt.Text = DevfrmCCtas
'            If (Modo = 4) And clien <> "" Then 'si insertar antes estaba lo de abajo
            If (Modo = 3 Or Modo = 4) And clien <> "" Then 'si insertar o modificar
                cad = cad & "  ¿Desea crearla?"
                If MsgBox(cad, vbYesNo) = vbYes Then
                    If InStr(1, Txt.Tag, "clientes") > 0 Then
                        InsertarCuentaCble DevfrmCCtas, clien
                    ElseIf InStr(1, Txt.Tag, "proveedor") > 0 Then
                        InsertarCuentaCble DevfrmCCtas, "", clien
                    ElseIf InStr(1, Txt.Tag, "agencias") > 0 Then
                        InsertarCuentaCble DevfrmCCtas, "", "", clien
                    End If
                    PonerNombreCuenta = clien
                End If
            Else
                MsgBox cad, vbExclamation
            End If
        Else
            Txt.Text = DevfrmCCtas
            PonerNombreCuenta = cad
        End If
    Else
        MsgBox cad, vbExclamation
'        Txt.Text = ""
        PonerNombreCuenta = ""
'        PonerFoco Txt
    End If
    DevfrmCCtas = ""

End Function




'DAVID: Cuentas del la Contabilidad
Public Function CuentaCorrectaUltimoNivel(ByRef Cuenta As String, ByRef devuelve As String) As Boolean
    'Comprueba si es numerica
    Dim sql As String
    Dim otroCampo As String
    
    CuentaCorrectaUltimoNivel = False
    If Cuenta = "" Then
        devuelve = "Cuenta vacia"
        Exit Function
    End If

    If Not IsNumeric(Cuenta) Then
        devuelve = "La cuenta debe de ser numérica: " & Cuenta
        Exit Function
    End If

    'Rellenamos si procede
    Cuenta = RellenaCodigoCuenta(Cuenta)

    '==========
    If Not EsCuentaUltimoNivel(Cuenta) Then
        devuelve = "No es cuenta de último nivel: " & Cuenta
        Exit Function
    End If
    '==================

    otroCampo = "apudirec"
    'BD 2: conexion a BD Conta
    sql = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Cuenta, "T", otroCampo)
    If sql = "" Then
        devuelve = "No existe la cuenta : " & Cuenta
        CuentaCorrectaUltimoNivel = True
        Exit Function
    End If

    'Llegados aqui, si que existe la cuenta
    If otroCampo = "S" Then 'Si es apunte directo
        CuentaCorrectaUltimoNivel = True
        devuelve = sql
    Else
        devuelve = "No es apunte directo: " & Cuenta
    End If
End Function


'DAVID
Public Function RellenaCodigoCuenta(vcodigo As String) As String
'Rellena con ceros hasta poner una cuenta.
'Ejemplo: 43.1 --> 430000001
Dim i As Integer
Dim J As Integer
Dim Cont As Integer
Dim cad As String

    RellenaCodigoCuenta = vcodigo
    If Len(vcodigo) > vEmpresa.DigitosUltimoNivel Then Exit Function
    
    i = 0: Cont = 0
    Do
        i = i + 1
        i = InStr(i, vcodigo, ".")
        If i > 0 Then
            If Cont > 0 Then Cont = 1000
            Cont = Cont + i
        End If
    Loop Until i = 0

    'Habia mas de un punto
    If Cont > 1000 Or Cont = 0 Then Exit Function

    'Cambiamos el punto por 0's  .-Utilizo la variable maximocaracteres, para no tener k definir mas
    i = Len(vcodigo) - 1 'el punto lo quito
    J = vEmpresa.DigitosUltimoNivel - i
    cad = ""
    For i = 1 To J
        cad = cad & "0"
    Next i

    cad = Mid(vcodigo, 1, Cont - 1) & cad
    cad = cad & Mid(vcodigo, Cont + 1)
    RellenaCodigoCuenta = cad
End Function

'DAVID
Public Function EsCuentaUltimoNivel(Cuenta As String) As Boolean
    EsCuentaUltimoNivel = (Len(Cuenta) = vEmpresa.DigitosUltimoNivel)
End Function

' ### [Monica] 07/09/2006
' copia de la gestion
Private Function InsertarCuentaCble(Cuenta As String, cadClien As String, Optional cadProve As String, Optional cadTrans As String) As Boolean
Dim sql As String
Dim vCliente As CCliente
Dim vProveedor As CProveedor
Dim vTranspor As CTransportista
Dim b As Boolean
Dim NombPais As String
Dim CtaBancoPropio As String
Dim vIban As String

    On Error GoTo EInsCta
    
    
    If vParamAplic.ContabilidadNueva Then
        sql = "INSERT INTO cuentas (codmacta,nommacta,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos,maidatos,webdatos,obsdatos,codpais, forpa, ctabanco"
    Else
        '[Monica]21/11/2014: añadida la forma de pago
        sql = "INSERT INTO cuentas (codmacta,nommacta,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos,maidatos,webdatos,obsdatos,pais, entidad, oficina, cc, cuentaba, forpa, ctabanco"
    End If
    
    '[Monica]22/11/2013: tema iban
    If vEmpresa.HayNorma19_34Nueva = 1 Then
        sql = sql & ", iban) "
    Else
        sql = sql & ") "
    End If
        
    sql = sql & " VALUES (" & DBSet(Cuenta, "T") & ","
    
    
    If cadClien <> "" Then
        Set vCliente = New CCliente
        If vCliente.LeerDatos(cadClien) Then
            '++[Monica] 10/12/2009 : En el cliente sí que tenemos el código del país
            NombPais = "ESPAÑA"
            If vCliente.CodPais <> 0 Then
                NombPais = DevuelveDesdeBDNew(cAgro, "paises", "letraspais", "codpaise", vCliente.CodPais, "N")
                If NombPais <> "" Then
                    '[Monica]20/11/2013: Si el cliente es de España no hay que concatenar ES España sino solo España, añado también el trim
                    If NombPais = "ES" Then NombPais = ""
                    
                    NombPais = Trim(NombPais & " " & DevuelveDesdeBDNew(cAgro, "paises", "nompaise", "codpaise", vCliente.CodPais, "N"))
                End If
            End If
            '++
        
            If Not vParamAplic.ContabilidadNueva Then
                sql = sql & DBSet(vCliente.Nombre, "T") & ",'S',1," & DBSet(vCliente.Nombre, "T") & "," & DBSet(vCliente.Domicilio, "T") & ","
                sql = sql & DBSet(vCliente.CPostal, "T") & "," & DBSet(vCliente.Poblacion, "T") & "," & DBSet(vCliente.Provincia, "T") & "," & DBSet(vCliente.NIF, "T") & "," & DBSet(vCliente.EMailAdm, "T") & "," & DBSet(vCliente.WebClien, "T") & "," & ValorNulo & "," & DBSet(NombPais, "T") & "," & DBSet(vCliente.Banco, "T", "S") & "," & DBSet(vCliente.Sucursal, "T", "S") & "," & DBSet(vCliente.DigControl, "T", "S") & "," & DBSet(vCliente.CuentaBan, "T", "S") & "," & DBSet(vCliente.ForPago, "N")
                
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    sql = sql & "," & ValorNulo & "," & DBSet(vCliente.Iban, "T", "S") & ")"
                Else
                    sql = sql & "," & ValorNulo & ")"
                End If
            Else
                sql = sql & DBSet(vCliente.Nombre, "T") & ",'S',1," & DBSet(vCliente.Nombre, "T") & "," & DBSet(vCliente.Domicilio, "T") & ","
                sql = sql & DBSet(vCliente.CPostal, "T") & "," & DBSet(vCliente.Poblacion, "T") & "," & DBSet(vCliente.Provincia, "T") & "," & DBSet(vCliente.NIF, "T") & "," & DBSet(vCliente.EMailAdm, "T") & "," & DBSet(vCliente.WebClien, "T") & "," & ValorNulo & "," & DBSet(NombPais, "T") & "," & DBSet(vCliente.ForPago, "N")
                
                vIban = MiFormat(vCliente.Iban, "") & MiFormat(vCliente.Banco, "0000") & MiFormat(vCliente.Sucursal, "0000") & MiFormat(vCliente.DigControl, "00") & MiFormat(vCliente.CuentaBan, "0000000000")
                
                sql = sql & "," & ValorNulo & "," & DBSet(vIban, "T", "S") & ")"
            End If
            
            ConnConta.Execute sql
            cadClien = vCliente.Nombre
            b = True
        Else
            b = False
        End If
        Set vCliente = Nothing
    End If
    
    If cadProve <> "" Then
        Set vProveedor = New CProveedor
        If vProveedor.LeerDatos(cadProve) Then
            If Not vParamAplic.ContabilidadNueva Then
                sql = sql & DBSet(vProveedor.Nombre, "T") & ",'S',1," & DBSet(vProveedor.NomComer, "T") & "," & DBSet(vProveedor.Domicilio, "T") & ","
                sql = sql & DBSet(vProveedor.CPostal, "T") & "," & DBSet(vProveedor.Poblacion, "T") & "," & DBSet(vProveedor.Provincia, "T") & "," & DBSet(vProveedor.NIF, "T") & "," & DBSet(vProveedor.EMailAdmon, "T") & "," & DBSet(vProveedor.WebProve, "T") & "," & ValorNulo & ",'ESPAÑA'," & DBSet(vProveedor.Banco, "T", "S") & "," & DBSet(vProveedor.Sucursal, "T", "S") & "," & DBSet(vProveedor.DigControl, "T", "S") & "," & DBSet(vProveedor.CuentaBan, "T", "S") & "," & DBSet(vProveedor.ForPago, "N")
                
                '[Monica]28/10/2016: faltaba la cta contable del banco propio
                CtaBancoPropio = ""
                If ComprobarCero(vProveedor.BancoPropio) <> 0 Then
                    CtaBancoPropio = DevuelveValor("select codmacta from banpropi where codbanpr = " & DBSet(vProveedor.BancoPropio, "N"))
                End If
            
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    sql = sql & "," & DBSet(CtaBancoPropio, "T", "S") & "," & DBSet(vProveedor.Iban, "T", "S") & ")"
                Else
                    sql = sql & "," & DBSet(CtaBancoPropio, "T", "S") & ")"
                End If
            Else
                sql = sql & DBSet(vProveedor.Nombre, "T") & ",'S',1," & DBSet(vProveedor.NomComer, "T") & "," & DBSet(vProveedor.Domicilio, "T") & ","
                sql = sql & DBSet(vProveedor.CPostal, "T") & "," & DBSet(vProveedor.Poblacion, "T") & "," & DBSet(vProveedor.Provincia, "T") & "," & DBSet(vProveedor.NIF, "T") & "," & DBSet(vProveedor.EMailAdmon, "T") & "," & DBSet(vProveedor.WebProve, "T") & "," & ValorNulo & ",'ESPAÑA'," & DBSet(vProveedor.ForPago, "N")
                
                '[Monica]28/10/2016: faltaba la cta contable del banco propio
                CtaBancoPropio = ""
                If ComprobarCero(vProveedor.BancoPropio) <> 0 Then
                    CtaBancoPropio = DevuelveValor("select codmacta from banpropi where codbanpr = " & DBSet(vProveedor.BancoPropio, "N"))
                End If
            
                vIban = MiFormat(vProveedor.Iban, "") & MiFormat(vProveedor.Banco, "0000") & MiFormat(vProveedor.Sucursal, "0000") & MiFormat(vProveedor.DigControl, "00") & MiFormat(vProveedor.CuentaBan, "0000000000")
            
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    sql = sql & "," & DBSet(CtaBancoPropio, "T", "S") & "," & DBSet(vIban, "T", "S") & ")"
                Else
                    sql = sql & "," & DBSet(CtaBancoPropio, "T", "S") & ")"
                End If
            
            End If
            
            ConnConta.Execute sql
            cadProve = vProveedor.Nombre
            b = True
        Else
            b = False
        End If
        Set vProveedor = Nothing
    End If
    
    If cadTrans <> "" Then
        Set vTranspor = New CTransportista
        If vTranspor.LeerDatos(cadTrans) Then
            If Not vParamAplic.ContabilidadNueva Then
                sql = sql & DBSet(vTranspor.Nombre, "T") & ",'S',1," & DBSet(Cuenta, "T") & "," & DBSet(vTranspor.Domicilio, "T") & ","
                sql = sql & DBSet(vTranspor.CPostal, "T") & "," & DBSet(vTranspor.Poblacion, "T") & "," & DBSet(vTranspor.Provincia, "T") & "," & DBSet(vTranspor.NIF, "T") & "," & DBSet(vTranspor.EMailAdmon, "T") & "," & DBSet(vTranspor.WebTrans, "T") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(vTranspor.Banco, "T", "S") & "," & DBSet(vTranspor.Sucursal, "T", "S") & "," & DBSet(vTranspor.DigControl, "T", "S") & "," & DBSet(vTranspor.CuentaBan, "T", "S") & "," & DBSet(vTranspor.ForPago, "N")
                
                '[Monica]28/10/2016: faltaba la cta contable del banco propio
                CtaBancoPropio = ""
                If ComprobarCero(vTranspor.BancoPropio) <> 0 Then
                    CtaBancoPropio = DevuelveValor("select codmacta from banpropi where codbanpr = " & DBSet(vTranspor.BancoPropio, "N"))
                End If
                
                
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    sql = sql & "," & DBSet(CtaBancoPropio, "T", "S") & "," & DBSet(vTranspor.Iban, "T", "S") & ")"
                Else
                    sql = sql & "," & DBSet(CtaBancoPropio, "T", "S") & ")"
                End If
            
            Else
                sql = sql & DBSet(vTranspor.Nombre, "T") & ",'S',1," & DBSet(Cuenta, "T") & "," & DBSet(vTranspor.Domicilio, "T") & ","
                sql = sql & DBSet(vTranspor.CPostal, "T") & "," & DBSet(vTranspor.Poblacion, "T") & "," & DBSet(vTranspor.Provincia, "T") & "," & DBSet(vTranspor.NIF, "T") & "," & DBSet(vTranspor.EMailAdmon, "T") & "," & DBSet(vTranspor.WebTrans, "T") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(vTranspor.ForPago, "N")
                
                '[Monica]28/10/2016: faltaba la cta contable del banco propio
                CtaBancoPropio = ""
                If ComprobarCero(vProveedor.BancoPropio) <> 0 Then
                    CtaBancoPropio = DevuelveValor("select codmacta from banpropi where codbanpr = " & DBSet(vTranspor.BancoPropio, "N"))
                End If
                
                vIban = MiFormat(vTranspor.Iban, "") & MiFormat(vTranspor.Banco, "0000") & MiFormat(vTranspor.Sucursal, "0000") & MiFormat(vTranspor.DigControl, "00") & MiFormat(vTranspor.CuentaBan, "0000000000")
                
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    sql = sql & "," & DBSet(CtaBancoPropio, "T", "S") & "," & DBSet(vIban, "T") & ")"
                Else
                    sql = sql & "," & DBSet(CtaBancoPropio, "T", "S") & ")"
                End If
            
            End If
            
            ConnConta.Execute sql
            cadTrans = vTranspor.Nombre
            b = True
        Else
            b = False
        End If
        Set vTranspor = Nothing
    End If
    
EInsCta:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Description, "Insertando cuenta contable", Err.Description
    End If
    InsertarCuentaCble = b
End Function




'=============================================================================
'==========     CONCEPTOS
'=============================================================================
'LAURA
Public Function PonerNombreConcepto(ByRef Txt As TextBox) As String
'Obtener el nombre de un concepto
Dim codConce As String
Dim cad As String

     If Txt.Text = "" Then
         PonerNombreConcepto = ""
         Exit Function
    End If
    codConce = Txt.Text
    If ConceptoCorrecto(codConce, cad) Then
        Txt.Text = Format(codConce, "000")
        PonerNombreConcepto = cad
    Else
        MsgBox cad, vbExclamation
        Txt.Text = ""
        PonerNombreConcepto = ""
        PonerFoco Txt
    End If
End Function


'LAURA
Public Function ConceptoCorrecto(ByRef Concep As String, ByRef devuelve As String) As Boolean
    Dim sql As String
    
    ConceptoCorrecto = False
 
    'BD 2: conexion a BD Conta
    sql = DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", Concep, "N")
    If sql = "" Then
        devuelve = "No existe el concepto : " & Concep
        Exit Function
    Else
        devuelve = sql
        ConceptoCorrecto = True
    End If
End Function


Public Function CalcularIva(Importe As String, articulo As String) As Currency
'devuelve el iva del Importe
'Ej el 16% de 120 = 19.2
Dim vImp As Currency
Dim vIva As Currency
Dim vArt As Currency
Dim Codiva As String

Dim IvaArt As Integer
Dim iva As String
Dim impiva As Currency
On Error Resume Next

    Importe = ComprobarCero(Importe)
    articulo = ComprobarCero(articulo)
    
    Codiva = DevuelveDesdeBD("codigiva", "sartic", "codartic", articulo, "N")
    iva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Codiva, "N")
    
    vImp = CCur(Importe)
    vIva = CCur(iva)
    
    impiva = ((vImp * vIva) / 100)
    impiva = Round(impiva, 2)
    
    CalcularIva = CStr(impiva)
    If Err.Number <> 0 Then Err.Clear

End Function


Public Function CalcularBase(Importe As String, articulo As String) As Currency
'devuelve la base del Importe
'Ej el 16% de 120 = 120-19.2 = 100.8
Dim vImp As Currency
Dim vIva As Currency
Dim vArt As Currency
Dim Codiva As String

Dim IvaArt As Integer
Dim iva As String
Dim impiva As Currency
On Error Resume Next

    Importe = ComprobarCero(Importe)
    articulo = ComprobarCero(articulo)
    
    Codiva = DevuelveDesdeBD("codigiva", "sartic", "codartic", articulo, "N")
    iva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Codiva, "N")
    
    vImp = CCur(Importe)
    vIva = CCur(iva)
    
    impiva = Round2(Importe / (1 + (vIva / 100)), 2)
    
    CalcularBase = CStr(impiva)
    If Err.Number <> 0 Then Err.Clear

End Function


'MONICA: Cuentas del la Contabilidad
Public Function NombreCuentaCorrecta(ByRef Cuenta As String) As String
    'Comprueba si es numerica
    Dim sql As String
    Dim otroCampo As String
    
' ### [Monica] 27/10/2006 añadida la linea siguiente condicion vParamAplic.NumeroConta = 0
' para que no saque nada si no hay contabilidad
    If Cuenta = "" Or vParamAplic.NumeroConta = 0 Then
         NombreCuentaCorrecta = ""
         Exit Function
    End If
    
    NombreCuentaCorrecta = ""
    If Cuenta = "" Then
        MsgBox "Cuenta vacia", vbExclamation
        Exit Function
    End If

    If Not IsNumeric(Cuenta) Then
        MsgBox "La cuenta debe de ser numérica: " & Cuenta, vbExclamation
        Exit Function
    End If

    'BD 2: conexion a BD Conta
    sql = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Cuenta, "T")
    If sql = "" Then
        MsgBox "No existe la cuenta : " & Cuenta, vbExclamation
    Else
        NombreCuentaCorrecta = sql
    End If

End Function


Public Function HayCobrosPagosPendientes(vCodmacta As String) As Boolean
Dim sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Nregs As Long

    On Error GoTo eHayCobrosPagosPendientes

    If vParamAplic.ContabilidadNueva Then
        sql = "select count(*) from cobros where codmacta = " & DBSet(vCodmacta, "T")
        sql = sql & " and (codrem is null or codrem = 0) and (transfer is null or transfer = 0) "
    Else
        sql = "select count(*) from scobro where codmacta = " & DBSet(vCodmacta, "T")
        sql = sql & " and (codrem is null or codrem = 0) and (transfer is null or transfer = 0) "
    End If
    Set Rs = New ADODB.Recordset
    Rs.Open sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0).Value) <> 0 Then Nregs = DBLet(Rs.Fields(0).Value)
    End If
            
    If vParamAplic.ContabilidadNueva Then
        sql = "select count(*) from pagos where codmacta = " & DBSet(vCodmacta, "T")
        sql = sql & " and (nrodocum is null or nrodocum = 0)"
    Else
        sql = "select count(*) from spagop where ctaprove = " & DBSet(vCodmacta, "T")
        sql = sql & " and (transfer is null or transfer = 0)"
    End If
    Set Rs = Nothing
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0).Value) <> 0 Then Nregs = Nregs + DBLet(Rs.Fields(0).Value)
    End If
    Set Rs = Nothing
    
    HayCobrosPagosPendientes = (Nregs <> 0)
    Exit Function
    
eHayCobrosPagosPendientes:
    MuestraError Err.Number, "Hay Cobros Pagos Pendientes", Err.Description
End Function

Public Function ActualizarCobrosPagosPdtes(vCodmacta As String, vBanco As String, vSucur As String, vDigcon As String, vCta As String, vIban As String, vFPago As String) As Boolean
Dim Sql2 As String
Dim vvIban As String
    
    On Error GoTo eActualizarCobrosPagosPdtes
    
    ConnConta.BeginTrans
    
    ActualizarCobrosPagosPdtes = False
    
    If vParamAplic.ContabilidadNueva Then
        vvIban = MiFormat(vIban, "") & MiFormat(vBanco, "0000") & MiFormat(vSucur, "0000") & MiFormat(vDigcon, "00") & MiFormat(vCta, "0000000000")
        
        Sql2 = "update cobros set iban = " & DBSet(vvIban, "T", "S")
    Else
        Sql2 = "update scobro set codbanco = " & DBSet(vBanco, "N", "S") & ", codsucur = " & DBSet(vSucur, "N", "S")
        Sql2 = Sql2 & ", digcontr = " & DBSet(vDigcon, "T", "S") & ", cuentaba = " & DBSet(vCta, "T", "S")
    
        '[Monica]22/11/2013: tema iban
        If vEmpresa.HayNorma19_34Nueva = 1 Then
            Sql2 = Sql2 & ", iban = " & DBSet(vIban, "T", "S")
        End If
    End If
    
    '[Monica]26/03/2015: se modifica tambien la forma de pago
    Sql2 = Sql2 & ", codforpa = " & DBSet(vFPago, "N")
    
    Sql2 = Sql2 & " where codmacta = " & DBSet(vCodmacta, "T")
    Sql2 = Sql2 & " and (codrem is null or codrem = 0) and (transfer is null or transfer = 0)"
    
    ConnConta.Execute Sql2
    
    
    If vParamAplic.ContabilidadNueva Then
        vvIban = MiFormat(vIban, "") & MiFormat(vBanco, "0000") & MiFormat(vSucur, "0000") & MiFormat(vDigcon, "00") & MiFormat(vCta, "0000000000")
        
        Sql2 = "update pagos set iban = " & DBSet(vvIban, "T", "S")
        
        '[Monica]26/03/2015: se modifica tambien la forma de pago
        Sql2 = Sql2 & ", codforpa = " & DBSet(vFPago, "N")
        
        
        Sql2 = Sql2 & " where codmacta = " & DBSet(vCodmacta, "T")
        Sql2 = Sql2 & " and (nrodocum is null or nrodocum = 0)"
    
    Else
        Sql2 = "update spagop set entidad = " & DBSet(vBanco, "T", "S") & ", oficina = " & DBSet(vSucur, "T", "S")
        Sql2 = Sql2 & ", cc = " & DBSet(vDigcon, "T", "S") & ", cuentaba = " & DBSet(vCta, "T", "S")
        
        '[Monica]22/11/2013: tema iban
        If vEmpresa.HayNorma19_34Nueva = 1 Then
            Sql2 = Sql2 & ", iban = " & DBSet(vIban, "T", "S")
        End If
        '[Monica]26/03/2015: se modifica tambien la forma de pago
        Sql2 = Sql2 & ", codforpa = " & DBSet(vFPago, "N")
        
        
        Sql2 = Sql2 & " where ctaprove = " & DBSet(vCodmacta, "T")
        Sql2 = Sql2 & " and (transfer is null or transfer = 0)"
    End If
   
    ConnConta.Execute Sql2
    
    ActualizarCobrosPagosPdtes = True
    ConnConta.CommitTrans
    Exit Function
    
eActualizarCobrosPagosPdtes:
    ConnConta.RollbackTrans
    MuestraError Err.Number, "Actualizar Cobros Pagos Pendientes", Err.Description
End Function






