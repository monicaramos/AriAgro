Attribute VB_Name = "ModFacturar"
Option Explicit

'===================================================================================
'Modulo para el traspaso de registros de cabecera y lineas de las tablas de ALBARAN
'A las tablas del FACTURACION
' o para pasar de las tablas de Mantenimientos a tablas de FACTURACION
'====================================================================================

'operador del albaran para facturas de Mantenimientos
Private OpeFactu As String
Private MesFactu As String 'mes a facturar para Mantenimientos
Private TipCoMan As String 'tipo de contrato del mantenimiento

'Variables comunes en Albaranes para la cabecera de la FACTURA
Private letraser As String

Private TipoAlb As String
Private TipoFac As String

'Variable con la WHERE que selecciona todos los Albaranes que forma parte de la Factura
Private cadW As String


Dim Errores As String
Dim ErroresAux As String





Private Sub AnyadirAvisos(Donde As String)
    Errores = Errores & vbCrLf & vbCrLf & Donde & vbCrLf
End Sub



Private Sub MostrarAvisos()
    frmMensajes.vCampos = Errores
    frmMensajes.OpcionMensaje = 13
    frmMensajes.Show vbModal
End Sub


'========================================================

Public Function ComprobarFechaVenci(FechaVenci As Date, Dia1 As Byte, Dia2 As Byte, Dia3 As Byte) As Date
Dim newFecha As Date
Dim b As Boolean

'=== Modificada Laura: 23/01/2007
    On Error GoTo ErrObtFec
    b = False
    
    '--- comprobar que tiene dias de pago para obtener nueva fecha
    If Not (Dia1 > 0 Or Dia2 > 0 Or Dia3 > 0) Then
        'si no tiene dias de pago la fecha es OK y fin
        ComprobarFechaVenci = FechaVenci
        Exit Function
    End If
        
    
    '--- Obtener nueva fecha del vencimiento
    newFecha = FechaVenci
    
    Do
        'si dia de la fecha vencimiento es uno de los 3 dias de pagos fecha es OK
        If Day(newFecha) = Dia1 Or Day(newFecha) = Dia2 Or Day(newFecha) = Dia3 Then
'            newFecha = CStr(newFecha)
            b = True
        Else
            'mientras esta en el mismo mes vamos aumentando dias hasta encontrar un dia de pago
            newFecha = DateAdd("d", 1, CDate(newFecha))
        End If
    Loop Until b = True Or Year(newFecha) = Year(FechaVenci) + 3
    
    ComprobarFechaVenci = newFecha
    Exit Function
    
ErrObtFec:
    MuestraError Err.Number, "Obtener Fecha vencimiento según dias de pago.", Err.Description
End Function





Public Function ComprobarFechaVenci_old(FechaVenci As Date, Dia1 As Byte, Dia2 As Byte, Dia3 As Byte) As Date
Dim fechaV As Date
'Dim cadDias As String
Dim F As String

    fechaV = FechaVenci
    If Dia1 <> 0 Or Dia2 <> 0 Or Dia3 <> 0 Then
        OrdenarDias Dia1, Dia2, Dia3
        If Dia1 >= Day(fechaV) Then
            fechaV = Format(Dia1 & "/" & Month(fechaV) & "/" & Year(fechaV), "dd/mm/yyyy")
        Else
            If Dia2 >= Day(fechaV) Then
                fechaV = Format(Dia2 & "/" & Month(fechaV) & "/" & Year(fechaV), "dd/mm/yyyy")
            Else
                If Dia3 >= Day(fechaV) Then
                    fechaV = Format(Dia3 & "/" & Month(fechaV) & "/" & Year(fechaV), "dd/mm/yyyy")
                
                Else
                    'coger el primero del mes siguiente
                    If Dia1 <> 0 Then
                        F = Dia1 & "/"
                        
                    ElseIf Dia2 <> 0 Then
                        F = Dia2 & "/"
'                        fechaV = Format(Dia2 & "/" & Month(fechaV) + 1 & "/" & Year(fechaV), "dd/mm/yyyy")
                    ElseIf Dia3 <> 0 Then
                        F = Dia3 & "/"
'                        fechaV = Format(Dia3 & "/" & Month(fechaV) + 1 & "/" & Year(fechaV), "dd/mm/yyyy")
                    End If
                    If Month(fechaV) + 1 < 13 Then
                        F = F & Month(fechaV) + 1 & "/" & Year(fechaV)
                    Else
                        F = F & "01/" & Year(fechaV) + 1
                    End If
                    fechaV = Format(F, "dd/mm/yyyy")
                End If
            End If
        End If

    End If
    ComprobarFechaVenci_old = fechaV
End Function





Private Sub OrdenarDias(Dia1 As Byte, Dia2 As Byte, Dia3 As Byte)
'Entran los dias desordenados: dia1=10, dia2=5, dia3=20
'devuelve los dias ordenados: dia1=5, dia2=10, dia3=20
Dim diaAux As Byte

    On Error GoTo EOrdenar

    If Dia1 < Dia2 And Dia1 < Dia3 Then
        'dia 1 es el menor
        If Dia2 > Dia3 Then
            diaAux = Dia2
            Dia2 = Dia3
            Dia3 = diaAux
        End If
    ElseIf Dia2 < Dia3 Then
        'dia2 es el menor
        diaAux = Dia1
        Dia1 = Dia2
        If diaAux < Dia3 Then
            Dia2 = diaAux
        Else
            Dia2 = Dia3
            Dia3 = diaAux
        End If
    Else
        'dia3 es el menor
        diaAux = Dia1
        Dia1 = Dia3
        If diaAux < Dia2 Then
            Dia3 = Dia2
            Dia2 = diaAux
        Else
            Dia3 = diaAux
        End If
    End If

EOrdenar:
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Function ComprobarMesNoGira(FecVenci As Date, MesNG As Byte, DiaVtoAt As Byte, Dia1 As Byte, Dia2 As Byte, Dia3 As Byte) As Date
Dim F As String

    If Month(FecVenci) = MesNG Then
        If DiaVtoAt > 0 Then
            F = DiaVtoAt & "/"
        Else
            F = Day(FecVenci) & "/"
        End If
        
        If Month(FecVenci) + 1 < 13 Then
            F = F & Month(FecVenci) + 1 & "/" & Year(FecVenci)
        Else
            F = F & "01/" & Year(FecVenci) + 1
        End If
        FecVenci = Format(F, "dd/mm/yyyy")
    End If
    ComprobarMesNoGira = FecVenci
End Function


Public Sub ImprimirFacturas(listaF As String, fechaF As String, Optional Sql As String, Optional FormatoFacturaTPV As Boolean)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NombreTabla As String

    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    NombreTabla = "facturas"

    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 12 'Facturas Clientes
    
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then
        Exit Sub
    End If

    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu


    If Sql <> "" Then
        'Llamo desde el menu de Reimprimir facturas y tengo construida la
        'cadena de seleccion D/H tipoMov, D/H NumFactu, D/H fecfactu
        cadSelect = Sql
        cadFormula = listaF
        cadParam = cadParam & fechaF
        numParam = numParam + 1
    Else
        'Llama desde PasarAlbaranes a  Facturas y al terminar las imprime
        '===================================================
        '================= FORMULA =========================
        'Cadena para seleccion Nº de Factura
        '---------------------------------------------------
        'Cod Tipo Movimiento
        devuelve = "({" & NombreTabla & ".codtipom}='" & TipoFac & "') "
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
        'Nº Factura
        devuelve = "({" & NombreTabla & ".numfactu} IN [" & listaF & "])"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
        'fecha factu
        devuelve = "(year({" & NombreTabla & ".fecfactu}) = " & Year(fechaF) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub

        cadSelect = cadFormula

        cadSelect = Replace(cadSelect, "[", "(")
        cadSelect = Replace(cadSelect, "]", ")")
    End If
    
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub

     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Factura Venta Cliente"
            .ConSubInforme = True
            .Show vbModal
    End With
End Sub


Public Function TraspasoMtosAFacturas(cadSQL As String, cadSel As String, FechaFact As String, OpeFact As String, banPr As String, MesFact As String, ByRef Lbl As label) As Boolean       'Fecha de la factura, Operador
'IN -> cadSQL: cadena para seleccion de los mantenimientos que vamos a Facturar
'      FechaFact: Fecha de la Factura
'      OpeFact: Operador Factura

'Desde Mantenimientos Genera las Facturas correspondientes
Dim RSmto As ADODB.Recordset 'Ordenados por: clien,dpto,forma pago, dtoppago, dtognral
Dim b As Boolean
Dim Sql As String

Dim vClien As CCliente 'aqui cargamos los datos del cliente del mantenimiento para grabar en scafac
Dim vFactu As CFactura

Dim ListFactu As String
Dim Conta2 As Long

    On Error GoTo ETraspasoMtoFac


    TraspasoMtosAFacturas = False

    'comprobamos que no haya nadie facturando
    DesBloqueoManual ("VENFAC") 'facturas de mantenimiento
    If Not BloqueoManual("VENFAC", "1") Then
        MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    'Bloqueamos todos los mantenimientos que vamos a facturar (cabeceras y lineas)
'    SQL = " (scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien ) INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom AND scaalb.numalbar=slialb.numalbar "
    Sql = " scaman "

    If Not BloqueaRegistro(Sql, cadSel) Then
        Screen.MousePointer = vbDefault
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("VENFAC")
        Exit Function
    End If






    'EMPEZAMOS LA FACTURA
    Set vFactu = New CFactura
    vFactu.FecFactu = FechaFact 'Fecha para las Facturas

    'Cuenta Prevista de Cobro de las Facturas
    vFactu.BancoPr = banPr
    vFactu.CuentaPrev = DevuelveDesdeBDNew(cAgro, "sbanpr", "codmacta", "codbanpr", banPr, "N")

    OpeFactu = OpeFact 'operador de la factura de mantenimiento
    MesFactu = MesFact 'mes a factura para los mantenimientos

    b = True

    'Marcar Mantenimientos que se van a Facturar
    '----------------------------------------

    Sql = cadSQL & " ORDER BY scaman.codclien, scaman.coddirec, scaman.nummante "
    Set RSmto = New ADODB.Recordset
    Conta2 = InStr(1, cadSQL, " FROM ")
    ListFactu = "Select count(*) " & Mid(cadSQL, Conta2)



    RSmto.Open ListFactu, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Lbl.Tag = RSmto.Fields(0)
    RSmto.Close



    Conta2 = 0
    ListFactu = ""
    RSmto.Open Sql, conn, adOpenKeyset, adLockPessimistic, adCmdText
    'Le pongo                KEYSET      pq quiero contar los registros
    'Cada MAntenimiento genera una factura
    'Calcular y Grabar Factura en las Tablas de Facturas
    '---    -------------------------------------------------
     While Not RSmto.EOF

           Conta2 = Conta2 + 1
           Lbl.Caption = Conta2 & " de " & Lbl.Tag
           Lbl.Refresh

            If (RSmto.RecordCount Mod 10) = 9 Then DoEvents
        'para cada mantenimiento de la tabla scaman seleccionado para facturar
        vFactu.BrutoFac = CCur(RSmto!Importe)
        'tipo de contrato del mantenimientos
        TipCoMan = RSmto!Codtipco

        'Datos de la Cabecera: Insertar en scafac
        '-----------------------------------------
        Set vClien = New CCliente
        If vClien.LeerDatos(RSmto!CodClien) Then
            'Datos cliente
            vFactu.Cliente = RSmto!CodClien
            vFactu.NombreClien = vClien.Nombre
            vFactu.DomicilioClien = vClien.Domicilio
            vFactu.CPostal = vClien.CPostal
            vFactu.Poblacion = vClien.Poblacion
            vFactu.Provincia = vClien.Provincia
            vFactu.NIF = vClien.NIF
            vFactu.Telefono = vClien.TfnoClien
            vFactu.DirDpto = DBLet(RSmto!CodDirec, "T")
            vFactu.NombreDirDpto = DBLet(RSmto!nomdirec, "T")
'            vFactu.Agente = vClien.Agente
            'forma de pago del mantenimiento
            vFactu.ForPago = RSmto!Codforpa
            vFactu.TipForPago = DevuelveDesdeBDNew(cAgro, "sforpa", "tipforpa", "codforpa", RSmto!Codforpa, "N")

            vFactu.DtoGnral = 0
            vFactu.DtoPPago = 0
            vFactu.Banco = DBLet(vClien.Banco, "N")
            vFactu.Sucursal = DBLet(vClien.Sucursal, "N")
            vFactu.Digcontrol = DBLet(vClien.Digcontrol, "T")
            vFactu.CuentaBan = DBLet(vClien.CuentaBan, "T")

            vFactu.Observacion = DBLet(RSmto!concefac, "T")




            If Not vFactu.PasarMtosAFactura(TipCoMan, OpeFactu, MesFactu, RSmto!numMante) Then
                If b Then b = False
            Else
'--monica
'                vClien.ActualizaUltFecMovim (FechaFact)


                'añadirlo a la lista de facturas a imprimir
                If ListFactu = "" Then
                    ListFactu = vFactu.NumFactu
                Else
                    ListFactu = ListFactu & "," & vFactu.NumFactu
                End If
            End If
        End If
        Set vClien = Nothing
        RSmto.MoveNext
    Wend

    RSmto.Close
    Set RSmto = Nothing

    Set vFactu = Nothing
    Lbl.Caption = "Finalizando proceso"
    Lbl.Refresh
    If b Then
        MsgBox "Las Facturas de los Mantenimientos seleccionados se generaron correctamente.", vbInformation
    Else
        Sql = "ATENCIÓN:" & vbCrLf
        MsgBox Sql & "No todas las Facturas se generaron correctamente!!!.", vbInformation
    End If

    'Desbloqueamos ya no estamos facturando
    DesBloqueoManual ("VENFAC")
    TerminaBloquear

    If ListFactu <> "" Then
        Lbl.Caption = "Imprimiend"
        Lbl.Refresh
        ImprimirFacturaMan 53, ListFactu, FechaFact
    End If


ETraspasoMtoFac:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Facturando Mantenimientos", Err.Description
    End If
End Function




Private Sub ImprimirFacturaMan(OpcionListado As Byte, ListFactu As String, FecFactu As String)
'Imprime una factura de Mantenimiento
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NombreTabla As String
    
    NombreTabla = "scafac"
    
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    If (OpcionListado = 53) Then indRPT = 12 'Facturas Clientes
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then
        Exit Sub
    End If
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Factura
    '---------------------------------------------------
    'Cod Tipo Movimiento
    devuelve = "{" & NombreTabla & ".codtipom}='FAM'"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    cadSelect = cadFormula
    
    'Nº Factura
    devuelve = "{" & NombreTabla & ".numfactu} IN [" & ListFactu & "]"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    devuelve = "{" & NombreTabla & ".numfactu} IN (" & ListFactu & ")"
    If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    
    'Fecha Factura
    devuelve = "year({" & NombreTabla & ".fecfactu})=" & Year(FecFactu)
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    'Fecha Factura en cadSelect
'        devuelve = "{" & NombreTabla & ".fecfactu}= '" & Format(FecFactu, FormatoFecha) & "'"
    If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    
   
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = OpcionListado
            .Titulo = ""
            .Show vbModal
    End With
End Sub














'Ventas de TICKET
'=================================================================
Public Function EliminarVenta(cadSQL As String) As Boolean
'Eliminamos de las tablas de ventas: scaven, sliven
Dim Sql As String

    On Error GoTo EElimVen

    EliminarVenta = False
    
    
    'ELiminar lineas venta
    Sql = "DELETE FROM sliven "
    Sql = Sql & " WHERE " & Replace(cadSQL, "scaven", "sliven")
    conn.Execute Sql
    
'    Espera 0.1
    
    'Eliminar Cabeceras venta
    Sql = "DELETE FROM scaven "
    Sql = Sql & " WHERE " & Replace(cadSQL, "sliven", "scaven")
    conn.Execute Sql
        
    EliminarVenta = True

EElimVen:
    If Err.Number <> 0 Then
        MsgBox Err.Number, "Eliminar venta.", Err.Description
        EliminarVenta = False
    Else
        EliminarVenta = True
    End If
End Function


Public Function TraspasoAlbaranesFacturas(cadSQL As String, cadWHERE As String, FechaFact As String, banPr As String, ByRef PBar1 As ProgressBar, ByRef LblBar As label, ImprimeLasFacturasGeneradas As Boolean, ByRef vTipoM As String, TextosCSB As String) As Boolean
'IN -> cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'      FechaFact: Fecha de la Factura
'      BanPr: Cod. de Banco Propio
'      Pbar1:  Una progressbar. Se puede mandar un NOTHING, y no pasa nada. Si no se manda
'              es que estamos en un proceso corto o que no necesitabaos un pb1, con lo cual NO muestro el PB1
'      Imprime: Si despues de generarlo los imprime
'
'       vTipom:  Que tipo de albaran es, para luego la impresion saber que factura imprime
'      TextosCSB:  Si lleva llevara 3 lineas para meter ent tesoreria

'Desde Albaranes Genera las Facturas correspondientes
Dim RsAlb As ADODB.Recordset 'Ordenados por: tipofac,clien,dpto,forma pago, dtoppago, dtognral
Dim b As Boolean
Dim Sql As String

'Aqui Guardamos los datos del Albaran Anterior para comparar con el actual
Dim antClien As Long
Dim antDirec As Long
Dim antForpa As Long
Dim antDtoPP As Single, antDtoGn As Single

'direc/dpto actual para controlar el valor nulo
Dim actDirec As Long

'Concatenamos todas las facturas generadas para listarlas en el informe
Dim ListFactu As String
Dim vFactuVta As CFacturaVta
Dim Inc As Integer
Dim condicion As Boolean 'condicion que comprueba para romper la agrupacion de albaranes a 1 factura

'Por si no mando una progressbar, que no de errores
Dim PgbVisible As Boolean

    On Error GoTo ETraspasoAlbFac

    TraspasoAlbaranesFacturas = False

    ListFactu = ""
        
    'comprobamos que no haya nadie facturando
    DesBloqueoManual ("VENFAC") 'facturas de venta
    If Not BloqueoManual("VENFAC", "1") Then
        MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    'Bloqueamos todos los albaranes que vamos a facturar (cabeceras y lineas)
    'Nota: esta bloqueando tambien los registros de la tabla clientes: sclien correspondientes
    Sql = " (scaalb INNER JOIN clientes ON scaalb.codclien=clientes.codclien ) INNER JOIN slialb ON scaalb.numalbar=slialb.numalbar "
    If Not BloqueaRegistro(Sql, cadWHERE) Then
        Screen.MousePointer = vbDefault
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("VENFAC")
        Exit Function
    End If
    
   
    'Inicializar la Progress Bar
    PgbVisible = False
    If Not (PBar1 Is Nothing) Then
        If PBar1.visible Then PgbVisible = True
    End If
    If PgbVisible Then
        If InStr(1, cadSQL, "clientes") Then
'            Sql = Replace(cadSQL, "scaalb.*, clientes.periodof", "count(*)") 'si hay INNER JOIN con clientes
            Sql = Replace(cadSQL, "scaalb.*, clientes.nomclien,clientes.domclien,clientes.codpobla,clientes.pobclien,clientes.proclien,clientes.cifclien,clientes.telclie1", "count(*)") 'si hay INNER JOIN con sclien
        Else
            Sql = Replace(cadSQL, "*", "count(*)") 'si NO hay INNER JOIN con sclien
        End If
        
        
        Set RsAlb = New ADODB.Recordset
        RsAlb.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RsAlb.EOF Then
            CargarProgresNew PBar1, CInt(RsAlb.Fields(0))
            LblBar.Caption = "Inicializando el proceso..."
        End If
        RsAlb.Close
        Set RsAlb = Nothing
    End If
    
        
    'EMPEZAMOS LA FACTURA
    Set vFactuVta = New CFacturaVta
    vFactuVta.FecFactu = FechaFact 'Fecha para las Facturas

    'Marcar Albaranes que se van a Facturar
    '----------------------------------------
    Sql = cadSQL & " ORDER BY scaalb.codclien,  scaalb.codforpa "
    Set RsAlb = New ADODB.Recordset
    RsAlb.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    'Agrupar los Albaranes posibles en una misma Factura
    'Calcular y Grabar Factura en la Tabla de Facturas
    'Albaran(scaalb, slialb) -> Factura (facturas, facturas_envases)
    '----------------------------------------------------
    'Agrupar albaranes en 1 factura por : codclien,codforpa
    b = True
    
    antClien = 0 'cliente
    antForpa = 0 'forma de pago
    
    cadW = ""
    Errores = ""
    Inc = 0
    
    While Not RsAlb.EOF
        TipoAlb = vParamAplic.CodTipomAlb '"ALV"
        Inc = Inc + 1
        
             
         LblBar.Caption = "Facturando: Facturas colectivas"
         
         condicion = (antClien <> RsAlb!CodClien) Or (antForpa <> RsAlb!Codforpa)
         
'             If (antClien <> RSalb!CodClien) Or (antDirec <> actDirec) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral) Then
         If condicion Then
         '-----
            If cadW <> "" Then 'Facturacion PEndiente
                cadW = cadW & ") "
                If Not vFactuVta.PasarAlbaranesAFactura2(TipoAlb, cadW, TextosCSB, ErroresAux, False) Then
                    If b Then b = False
                    AnyadirAvisos ErroresAux
                Else 'añadirlo a la lista de facturas a imprimir
                    If ListFactu = "" Then
                        ListFactu = vFactuVta.NumFactu
                    Else
                        ListFactu = ListFactu & "," & vFactuVta.NumFactu
                    End If
                End If
                If PgbVisible Then
                    LblBar.Caption = "Cliente: " & Format(vFactuVta.Cliente, "000000") & " " & vFactuVta.NombreClien
                    IncrementarProgresNew PBar1, Inc
                    Inc = 0
                End If
                espera 0.2
                
                'Empezamos una nueva Factura
                cadW = ""
            End If
            'Generar una Factura nueva
            vFactuVta.Cliente = RsAlb!CodClien
            vFactuVta.NombreClien = RsAlb!nomclien
            vFactuVta.DomicilioClien = DBLet(RsAlb!domclien, "T")
            vFactuVta.CPostal = DBLet(RsAlb!codPobla, "T")
            vFactuVta.Poblacion = DBLet(RsAlb!pobclien, "T")
            vFactuVta.Provincia = DBLet(RsAlb!proclien, "T")
            vFactuVta.NIF = DBLet(RsAlb!cifclien, "T")
            vFactuVta.Telefono = DBLet(RsAlb!telclie1, "T")
            vFactuVta.ForPago = RsAlb!Codforpa
            vFactuVta.TipForPago = DevuelveDesdeBDNew(cAgro, "forpago", "tipoforp", "codforpa", RsAlb!Codforpa, "N")
            cadW = "  scaalb.numalbar IN (" & RsAlb!numalbar
        Else
            cadW = cadW & ", " & RsAlb!numalbar
        End If
    
        'Guardamos datos del registro anterior
        antClien = RsAlb!CodClien
        antForpa = RsAlb!Codforpa
        RsAlb.MoveNext
    Wend
    RsAlb.Close
    Set RsAlb = Nothing
        
    'Facturar la ultima Factura generada del blucle
    If cadW <> "" Then
        cadW = cadW & ")"
        If PgbVisible Then LblBar.Caption = "Cliente: " & Format(vFactuVta.Cliente, "000000") & " - " & vFactuVta.NombreClien
        
        If Not vFactuVta.PasarAlbaranesAFactura2(TipoAlb, cadW, TextosCSB, ErroresAux, False) Then
            If b Then b = False
            AnyadirAvisos "Error Facturando el Cliente: " & Format(vFactuVta.Cliente, "000000") & " " & vFactuVta.NombreClien & vbCrLf & ErroresAux
        Else 'añadirlo a la lista de facturas a imprimir
            If ListFactu = "" Then
                ListFactu = vFactuVta.NumFactu
            Else
                ListFactu = ListFactu & "," & vFactuVta.NumFactu
            End If
        End If
        If PgbVisible Then
'            LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " - " & vFactu.NombreClien
            IncrementarProgresNew PBar1, Inc
        End If
        espera 0.2
    End If
    
    TipoFac = vFactuVta.CodTipoM
    Set vFactuVta = Nothing
    TraspasoAlbaranesFacturas = True
    
    If b Then
        LblBar.Caption = "Proceso finalizado correctamente."
        MsgBox "Las Facturas de los Albaranes seleccionados se generaron correctamente.", vbInformation
    Else
        LblBar.Caption = "Proceso finalizado con errores."
        Sql = "ATENCIÓN:" & vbCrLf
        MsgBox Sql & "No todas las Facturas se generaron correctamente!!!.", vbExclamation
        If Errores <> "" Then MostrarAvisos
    End If
    
    espera 0.2
    
    'Desbloqueamos ya no estamos facturando
    DesBloqueoManual ("VENFAC")
    TerminaBloquear
    
    
    If ImprimeLasFacturasGeneradas Then
        If ListFactu <> "" Then
            ImprimirFacturas ListFactu, FechaFact, , False
        End If
    End If
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Facturando Albaranes", Err.Description
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("VENFAC")
        TerminaBloquear
    End If
End Function
