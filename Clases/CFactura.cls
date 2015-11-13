VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'VENTAS
'FACTURA: se corresponde con la tabla ariges.scafac (cabecera de facturas)
'-------------------------------------------------------------------------

'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodTipoM As String 'Tipo de Movimiento="FAV", "FAR" (Factura venta o reparacion)
Private mNumFactu As Long 'Nº Factura
Private mFecFactu As String 'FEcha Factura
Private mCodClien As String 'Cod. cliente

Private mNomClien As String 'nombre del cliente
Private mDomClien As String 'domicilio
Private mCodpobla As String 'cpostal
Private mPobClien As String 'poblacion
Private mProClien As String 'provincia
Private mNIFClien As String 'NIF proveedor
Private mTelclien As String 'telefono

Private mCoddirec As String 'codigo direc./dpto
Private mNomdirec As String  'nombre direc./dpto
Private mCodAgent As Integer 'codigo de agente


Private mCodbanco As String 'cod. banco
Private mCodsucur As String 'cod. sucur
Private mDigContr As String 'digito control
Private mCuentaba As String 'cuenta bancaria
'[Monica]222/11/2013: Tema iban
Private mIban As String

Private mCodForpa As String 'Cod. Forma de pago
Private mTipForpa As Byte 'tipo de forma de pago: efectivo,...

Private mDtoPpago As Currency 'descuento pronto pago
Private mDtoGnral As Currency 'descuento general
Private mBrutofac As Currency 'importe bruto de la factura
Private mImpPPago As Currency 'Importe al aplicar al bruto el dto pronto pago
Private mImpGnral As Currency  'Importe al aplicar al bruto el dto general
Private mBaseImp As Currency  'Base imponible de la factura (bruto - dtopp - dtogn)

Private mBaseiva1 As Currency  'base imponible tipo IVA 1
Private mBaseiva2 As Currency  'base imponible tipo IVA 2
Private mBaseiva3 As Currency  'base imponible tipo IVA 3
Private mTipoiva1 As Byte   'Cod. tipo de iva 1
Private mTipoiva2 As Byte   'Cod. tipo de iva 2
Private mTipoiva3 As Byte   'Cod. tipo de iva 3
Private mPorciva1 As Currency    '% de iva 1
Private mPorciva2 As Currency    '% de iva 2
Private mPorciva3 As Currency    '% de iva 3

Private mImpoiva1 As Currency   'Importe de iva 1
Private mImpoiva2 As Currency   'Importe de iva 2
Private mImpoiva3 As Currency   'Importe de iva 3

Private mTotalfac As String 'total factura

Private mLetraSer As String 'Letra de serie para los cobros
Private mBanPr As String 'Banco propio que ponemos como cuenta prevista de cobro
Private mCtaPrev As String 'Cuenta prevista de cobro

Private mNumtermi As Integer 'terminal de venta del TPV
Private mNumventa As Long 'Nº de venta del TPV


'David.
'Para las facturas de mantenimeinto, hay un campo que es "concepto factura" varchar(60)
'Si tiene valor entonces ese valor es el que grabaremos en la linea y como obervacion de la factura
'meteremos lo de antes ("texto" de: mes de año.  Ejemplo: manteimiento de Junio de 2007
'Si no tiene valor seguira como estaba
Private mObservacion As String

' VBLES para los Mantenimientos -------------------------------------------
Dim TipCoMan As String 'tipo de contrato para Facturas de Manteniminetos
Dim OpeFactu As String 'Operador conectado o el que pasamos como parametro
Dim AlmFactu As String 'Cod Almacen del trabajador conectado, para insertar lines mantenim, desplaz, bonificaciones,....
Dim MesFactu As String
Dim numMante As String
Dim ArticFac As String 'cod articulo que se va a utilizar para facturar mantenimientos
'--------------------------------------------------------------------------

' VBLES para los Desplazamientos ----------------------------------
Dim TotalKm As Integer
Dim PrecioKm As Currency
Dim ImporteL As Currency
Dim Ampliacion As String  'ampliacion de la linea: "Dia,Dia,Dia..." de los albaranes con desplazamiento
Dim codArtic As String
Dim despIVA As Integer 'tipo de IVA del articulo para desplazamientos
'------------------------------------------------------------------

' VBLES para las Bonificaciones -----------------------------------
Dim TotBonif1 As Currency 'Total de bonificacion a factura con articulo de tipo de IVA1
Dim IVABonif1 As Byte 'tipo de IVA
Dim TotBonif2 As Currency
Dim IVABonif2 As Byte
Dim TotBonif3 As Currency
Dim IVABonif3 As Byte
Dim SQLBonif As String
'------------------------------------------------------------------

' VBLES para los Cheques Regalo -----------------------------------
Dim ImpCheque As Currency
'------------------------------------------------------------------




'vbles para las facturas que vienen de una venta de ticket ---------
Private NumTicket As String
Private NumAlbTicket As String '01/09/06 laura
'-------------------------------------------------------------------


'# Laura 14/11/2006 recuperar facturas Alzira
Private EstaRecuFac As Boolean


'------------------------------------------------
'Propiedades del modulo CFactura
'------------------------------------------------

'**** Tipo de Movimiento de la Factura

Public Property Let codTipoM(ByVal vData As String)
     mCodTipoM = vData
End Property

Public Property Get codTipoM() As String
     codTipoM = mCodTipoM
End Property


'**** Nº de la Factura

Public Property Let NumFactu(ByVal vData As Long)
     mNumFactu = vData
End Property

Public Property Get NumFactu() As Long
     NumFactu = mNumFactu
End Property


'**** Fecha de la Factura

Public Property Let FecFactu(ByVal vData As String)
     mFecFactu = vData
End Property

Public Property Get FecFactu() As String
     FecFactu = mFecFactu
End Property


'**** Cliente de la Factura

Public Property Let Cliente(ByVal vData As String)
     mCodClien = vData
End Property

Public Property Get Cliente() As String
     Cliente = mCodClien
End Property


'**** Nombre del Cliente de la Factura

Public Property Let NombreClien(ByVal vData As String)
     mNomClien = vData
End Property

Public Property Get NombreClien() As String
     NombreClien = mNomClien
End Property


'**** Domicilio del Cliente de la Factura

Public Property Let DomicilioClien(ByVal vData As String)
     mDomClien = vData
End Property

Public Property Get DomicilioClien() As String
     DomicilioClien = mDomClien
End Property


'**** CPostal del Cliente de la Factura

Public Property Let CPostal(ByVal vData As String)
     mCodpobla = vData
End Property

Public Property Get CPostal() As String
     CPostal = mCodpobla
End Property


'**** Poblacion del Cliente de la Factura

Public Property Let Poblacion(ByVal vData As String)
     mPobClien = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mPobClien
End Property


'**** Provincia del Cliente de la Factura

Public Property Let Provincia(ByVal vData As String)
     mProClien = vData
End Property

Public Property Get Provincia() As String
     Provincia = mProClien
End Property


'**** NIF del Cliente de la Factura

Public Property Let NIF(ByVal vData As String)
     mNIFClien = vData
End Property

Public Property Get NIF() As String
     NIF = mNIFClien
End Property


'**** Telefono del Cliente de la Factura

Public Property Let Telefono(ByVal vData As String)
     mTelclien = vData
End Property

Public Property Get Telefono() As String
     Telefono = mTelclien
End Property


'**** Direc./Dpto de Cliente de la Factura

Public Property Let DirDpto(ByVal vData As String)
     mCoddirec = vData
End Property

Public Property Get DirDpto() As String
     DirDpto = mCoddirec
End Property


'**** Nombre Direc./Dpto de Cliente de la Factura

Public Property Let NombreDirDpto(ByVal vData As String)
     mNomdirec = vData
End Property

Public Property Get NombreDirDpto() As String
     NombreDirDpto = mNomdirec
End Property


'**** Agente de Cliente de la Factura

Public Property Let Agente(ByVal vData As Integer)
     mCodAgent = vData
End Property

Public Property Get Agente() As Integer
     Agente = mCodAgent
End Property


'**** Forma Pago de la Factura

Public Property Let ForPago(ByVal vData As String)
     mCodForpa = vData
End Property

Public Property Get ForPago() As String
     ForPago = mCodForpa
End Property


'**** Tipo de Forma Pago de la Factura
'(para tickets saber si insertar en tesoreria o no: si es efectivo no se inserta en scobro)

Public Property Let TipForPago(ByVal vData As Byte)
     mTipForpa = vData
End Property

Public Property Get TipForPago() As Byte
     TipForPago = mTipForpa
End Property


'**** Cuenta Prevista de cobro de la Factura

Public Property Let LetraSerie(ByVal vData As String)
     mLetraSer = vData
End Property

Public Property Get LetraSerie() As String
     LetraSerie = mLetraSer
End Property


'**** Banco propio de pago de la Factura

Public Property Let BancoPr(ByVal vData As String)
     mBanPr = vData
End Property

Public Property Get BancoPr() As String
     BancoPr = mBanPr
End Property


'**** Cuenta Prevista de cobro de la Factura

Public Property Let CuentaPrev(ByVal vData As String)
     mCtaPrev = vData
End Property

Public Property Get CuentaPrev() As String
     CuentaPrev = mCtaPrev
End Property


'**** Banco de la Factura

Public Property Let Banco(ByVal vData As String)
     mCodbanco = vData
End Property

Public Property Get Banco() As String
     Banco = mCodbanco
End Property


'**** Sucursal de la Factura

Public Property Let Sucursal(ByVal vData As String)
     mCodsucur = vData
End Property

Public Property Get Sucursal() As String
     Sucursal = mCodsucur
End Property


'**** Digito Control de la Factura

Public Property Let DigControl(ByVal vData As String)
     mDigContr = vData
End Property

Public Property Get DigControl() As String
     DigControl = mDigContr
End Property


'**** Cuenta Bancaria de la Factura

Public Property Let CuentaBan(ByVal vData As String)
     mCuentaba = vData
End Property

Public Property Get CuentaBan() As String
     CuentaBan = mCuentaba
End Property


'**** Iban

Public Property Let Iban(ByVal vData As String)
     mIban = vData
End Property

Public Property Get Iban() As String
     Iban = mIban
End Property



'**** Descuento pronto pago de la Factura

Public Property Let DtoPPago(ByVal vData As Currency)
     mDtoPpago = vData
End Property

Public Property Get DtoPPago() As Currency
     DtoPPago = mDtoPpago
End Property


'**** Descuento general de la Factura

Public Property Let DtoGnral(ByVal vData As Currency)
     mDtoGnral = vData
End Property

Public Property Get DtoGnral() As Currency
     DtoGnral = mDtoGnral
End Property


'**** Importe Bruto Factura

Public Property Let BrutoFac(ByVal vData As Currency)
     mBrutofac = vData
End Property

Public Property Get BrutoFac() As Currency
     BrutoFac = mBrutofac
End Property


'**** Importe pronto pago de la Factura

Public Property Let ImpPPago(ByVal vData As Currency)
     vData = vData
     mImpPPago = vData
End Property

Public Property Get ImpPPago() As Currency
     ImpPPago = mImpPPago
End Property


'**** Importe general de la Factura

Public Property Let ImpGnral(ByVal vData As Currency)
     mImpGnral = vData
End Property

Public Property Get ImpGnral() As Currency
     ImpGnral = mImpGnral
End Property


'**** Importe Base Imponible Factura

Public Property Let BaseImp(ByVal vData As Currency)
     mBaseImp = vData
End Property

Public Property Get BaseImp() As Currency
     BaseImp = mBaseImp
End Property


'**** Base IVA1 de la Factura

Public Property Let BaseIVA1(ByVal vData As Currency)
     mBaseiva1 = vData
End Property

Public Property Get BaseIVA1() As Currency
     BaseIVA1 = mBaseiva1
End Property


'**** Base IVA2 de la Factura

Public Property Let BaseIVA2(ByVal vData As Currency)
     mBaseiva2 = vData
End Property

Public Property Get BaseIVA2() As Currency
     BaseIVA2 = mBaseiva2
End Property


'**** Base IVA3 de la Factura

Public Property Let BaseIVA3(ByVal vData As Currency)
     mBaseiva3 = vData
End Property

Public Property Get BaseIVA3() As Currency
     BaseIVA3 = mBaseiva3
End Property


'**** Tipo iva 1 de la Factura

Public Property Let TipoIVA1(ByVal vData As Byte)
     mTipoiva1 = vData
End Property

Public Property Get TipoIVA1() As Byte
     TipoIVA1 = mTipoiva1
End Property


'**** Tipo iva 2 de la Factura

Public Property Let TipoIVA2(ByVal vData As Byte)
     mTipoiva2 = vData
End Property

Public Property Get TipoIVA2() As Byte
     TipoIVA2 = mTipoiva2
End Property


'**** Tipo iva 3 de la Factura

Public Property Let TipoIVA3(ByVal vData As Byte)
     mTipoiva3 = vData
End Property

Public Property Get TipoIVA3() As Byte
     TipoIVA3 = mTipoiva3
End Property


'**** % de iva 1 de la Factura

Public Property Let PorceIVA1(ByVal vData As Currency)
     mPorciva1 = vData
End Property

Public Property Get PorceIVA1() As Currency
     PorceIVA1 = mPorciva1
End Property


'**** % de iva 2 de la Factura

Public Property Let PorceIVA2(ByVal vData As Currency)
     mPorciva2 = vData
End Property

Public Property Get PorceIVA2() As Currency
     PorceIVA2 = mPorciva2
End Property


'**** % de iva 3 de la Factura

Public Property Let PorceIVA3(ByVal vData As Currency)
     mPorciva3 = vData
End Property

Public Property Get PorceIVA3() As Currency
     PorceIVA3 = mPorciva3
End Property


'**** Importe de IVA 1 de la Factura

Public Property Let ImpIVA1(ByVal vData As Currency)
     mImpoiva1 = vData
End Property

Public Property Get ImpIVA1() As Currency
     ImpIVA1 = mImpoiva1
End Property


'**** Importe de IVA 2 de la Factura

Public Property Let ImpIVA2(ByVal vData As Currency)
     mImpoiva2 = vData
End Property

Public Property Get ImpIVA2() As Currency
     ImpIVA2 = mImpoiva2
End Property


'**** Importe de IVA 3 de la Factura

Public Property Let ImpIVA3(ByVal vData As Currency)
     mImpoiva3 = vData
End Property

Public Property Get ImpIVA3() As Currency
     ImpIVA3 = mImpoiva3
End Property


'**** Total Factura

Public Property Let TotalFac(ByVal vData As String)
     mTotalfac = vData
End Property

Public Property Get TotalFac() As String
     TotalFac = mTotalfac
End Property


'**** Nº terminal de la venta del TPV

Public Property Let NumTerminal(ByVal vData As Integer)
     mNumtermi = vData
End Property

Public Property Get NumTerminal() As Integer
     NumTerminal = mNumtermi
End Property



'**** Nº venta del TPV

Public Property Let NumVenta(ByVal vData As Long)
     mNumventa = vData
End Property

Public Property Get NumVenta() As Long
     NumVenta = mNumventa
End Property


'**** Para los MANTENIMIENTOS.  sera 1 la linea. Si esta vacio lo cargaremos con la amplicaicon normal

Public Property Let Observacion(ByVal vData As String)
     mObservacion = vData
End Property

Public Property Get Observacion() As String
     Observacion = mObservacion
End Property




'------------------------------------------------
'Procedimientos del modulo CFactura
'------------------------------------------------

Public Function LeerDatos(vCodtipom As String, vNumfactu As String, vFecfactu As String) As Boolean
'Leer los datos de una factura almacenada en la tabla scafac
'Lee de la BD: Ariges, Tabla: scafac
'OUT: True si lee los datos correctamente
Dim Rs As ADODB.Recordset
Dim Mens As String
Dim Sql As String

    On Error GoTo ELeer
    
    
    LeerDatos = False
    '[Monica]22/11/2013: Tema iban
    Sql = "SELECT codtipom,numfactu,fecfactu,codclien,codforpa,codbanco,codsucur,digcontr,cuentaba,brutofac, totalfac, iban FROM scafac "
    Sql = Sql & " WHERE codtipom='" & vCodtipom & "'"
    Sql = Sql & " AND numfactu=" & vNumfactu
    Sql = Sql & " AND fecfactu='" & Format(vFecfactu, FormatoFecha) & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If Rs.EOF Then
        LeerDatos = False
    Else
        Mens = "Datos Factura"
        mCodTipoM = Rs!codTipoM
        mNumFactu = CStr(Rs!NumFactu)
        mFecFactu = CStr(Rs!FecFactu)
        mCodClien = CStr(Rs!CodClien)
        ForPago = CStr(Rs!Codforpa)
        TipForPago = DevuelveDesdeBDNew(cAgro, "sforpa", "tipforpa", "codforpa", ForPago, "N")
        mCodbanco = CStr(DBLet(Rs!codbanco, "N"))
        mCodsucur = CStr(DBLet(Rs!codsucur, "N"))
        mDigContr = DBLet(Rs!digcontr, "T")
        mCuentaba = DBLet(Rs!cuentaba, "T")
        '[Monica]22/11/2013: Tema iban
        mIban = DBLet(Rs!Iban, "T")
        mBrutofac = CStr(Rs!BrutoFac)
        mTotalfac = CStr(Rs!TotalFac)

        mLetraSer = ObtenerLetraSerie(mCodTipoM)
        
        'si venimos de proceso de insertar tenemos la cuenta prevista de cobro que
        'se introdujo en el formulario de pasar albaranes a factura
        If CuentaPrev = "" Then
            'leer la cuenta prevista de cobros de la tabla de contabilidad: conta.scobro
            'cuando vamos a borrar una factura porque modificarmos
            'ya que este dato se pidio al facturar (SI EXISTE LA FACTURA EN TESORERIA)
            Sql = "SELECT COUNT(*) FROM scobro WHERE numserie='" & LetraSerie & "' and codfaccl=" & NumFactu
            Sql = Sql & " AND fecfaccl=" & DBSet(FecFactu, "F")
            If RegistrosAListar(Sql, cConta) > 0 Then
                Sql = DevuelveDesdeBDNew(cConta, "scobro", "ctabanc1", "numserie", mLetraSer, "T", , "codfaccl", vNumfactu, "N", "fecfaccl", vFecfactu, "F")
                If Sql <> "" Then
                    mCtaPrev = Sql
                    LeerDatos = True
                Else
                    LeerDatos = False
                    Mens = "La cuenta prevista de cobro para la factura no puede ser nula."
                End If
            Else
                LeerDatos = True
            End If
        Else
            LeerDatos = True
        End If
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set Rs = Nothing
    LeerDatos = False
End Function


'======================================================================
'GRABAR EN TESORERIA
'======================================================================

Public Function InsertarEnTesoreria(MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim Sql As String, Codmacta As String, textcsb33 As String
Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim CadValues3 As String
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single 'importe para insertar en la svenci
Dim ImpVenci2 As Single 'importe para insertar en conta.scobro
Dim i As Byte


    On Error GoTo EInsertarTesoreria

    b = False
    InsertarEnTesoreria = False
    CadValues = ""
    CadValues2 = ""

    'campo para insertar en conta.scobro de Tesoreria
    textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd,mm,yyyy") & "'"

    'Obtener el Nº de Vencimientos de la forma de pago
    Sql = "SELECT numerove, primerve, restoven FROM sforpa WHERE codforpa=" & ForPago
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 Then
            'Obtener los dias de pago del cliente
            Sql = " SELECT  diapago1, diapago2, diapago3, mesnogir, diavtoat, codmacta "
            Sql = Sql & " FROM sclien "
            Sql = Sql & " WHERE codclien=" & Cliente
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            Codmacta = DBSet(Rs!Codmacta, "T")
'            textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd,mm,yyyy") & "'"
            
            If Not Rs.EOF Then
                cadValuesAux = "('" & codTipoM & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "
                CadValuesAux2 = "('" & LetraSerie & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "
                '-------- Primer Vencimiento
                i = 1
                'FECHA VTO
                FecVenci = CDate(FecFactu)
                '=== Laura 23/01/2007
                'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                '===
                'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                If TipForPago <> 0 Then
                    FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                Else
                    FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                End If
                'Comprobar si cliente tiene mes a no girar
                FecVenci1 = FecVenci
                If CInt(DBLet(Rs!mesnogir, "N")) <> 0 Then
                    FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                End If
                
                'Comprobar si cliente tiene dia de vencimiento atrasado
                CadValues = cadValuesAux & i & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                CadValues2 = CadValuesAux2 & i & ", "
                CadValues2 = CadValues2 & Codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                
                'IMPORTE del Vencimiento
                If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFac
                    ImpVenci2 = TotalFac - ImpCheque
                Else
                    ImpVenci = Round2(TotalFac / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFac - ImpCheque) / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFac Then
                        ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * rsVenci!numerove), 2)
                    End If
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If (ImpVenci2 * rsVenci!numerove) + ImpCheque <> TotalFac Then
                        ImpVenci2 = Round(ImpVenci2 + (TotalFac - ImpCheque - (ImpVenci2 * rsVenci!numerove)), 2)
                    End If
                End If
                CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N", "S") & ", " & DBSet(Sucursal, "N", "S") & ", " & DBSet(DigControl, "T", "S") & ", " & DBSet(CuentaBan, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                'departamento
                CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") '& ")"
                
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    CadValues2 = CadValues2 & "," & DBSet(Iban, "T", "S") & ")"
                Else
                    CadValues2 = CadValues2 & ")"
                End If
                
                
                'Resto Vencimientos
                '--------------------------------------------------------------------
                For i = 2 To rsVenci!numerove
                   'FECHA Resto Vencimientos
                    '=== Laura 23/01/2007
                    'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                    FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    '===
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
                        FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    'Comprobar si cliente tiene mes a no girar
                    FecVenci1 = FecVenci
                    If DBLet(Rs!mesnogir, "N") <> "0" Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    End If
                        
                    CadValues = CadValues & ", " & cadValuesAux & i & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    CadValues2 = CadValues2 & ", " & CadValuesAux2 & i & ", " & Codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = Round2(TotalFac / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFac - ImpCheque) / rsVenci!numerove, 2)
                    CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                    CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", " & DBSet(Banco, "N", "S") & ", " & DBSet(Sucursal, "N", "S") & ", " & DBSet(DigControl, "T", "S") & ", " & DBSet(CuentaBan, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N") & ", "
                    CadValues2 = CadValues2 & DBSet(Me.DirDpto, "N", "S") '& ")"
                    
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        CadValues2 = CadValues2 & "," & DBSet(Iban, "T", "S") & ")"
                    Else
                        CadValues2 = CadValues2 & ")"
                    End If
                Next i
                
'--monica
'                '--- Cheque regalo: laura 1/12/2006
'                'si hay cheque regalo insertar una linea más para la forma de pago correspondiente y el importe del cheque
'                If ImpCheque > 0 Then
'                    CadValues2 = CadValues2 & ", " & CadValuesAux2 & i & "," & Codmacta & ", " & vParamAplic.ForPagoChequeRegalo & ", "
'
'                    'FECHA VTO
'                    CadValuesAux2 = "primerve"
'                    TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", vParamAplic.ForPagoChequeRegalo, "N", CadValuesAux2)
'                    FecVenci = CDate(FecFactu)
'                    FecVenci = FecVenci + CByte(CadValuesAux2)
'                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
'                    If TipForPago <> 0 Then
'            '            FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
'                        MsgBox "FALTA cheque regalo con forma de pago no en EFECTIVO", vbInformation
'                    Else
'                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
'                    End If
'
'                    CadValues2 = CadValues2 & DBSet(FecVenci, "F") & ", "
'                    CadValues2 = CadValues2 & DBSet(ImpCheque, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
'                    'departamento
'                    CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ")"
'                End If
            End If
        End If
        Rs.Close
        Set Rs = Nothing
    End If
    rsVenci.Close
    Set rsVenci = Nothing
    
    If CadValues <> "" Then
        Sql = "INSERT INTO svenci (codtipom, numfactu, fecfactu, ordefect, fecefect, impefect)"
        Sql = Sql & " VALUES " & CadValues
        Conn.Execute Sql
    End If
    
    
    'Grabar tabla scobro de la CONTABILIDAD
    '-------------------------------------------------
    If CadValues2 <> "" Then
        '01/09/06
'        If (NumTicket = "") Or (NumTicket <> "" And TipForPago <> 0) Then
            If CuentaPrev <> "" Then
                'antes de grabar en la scobro comprobar que existe en conta.sforpa la
                'forma de pago de la factura. Sino existe insertarla
                'vemos si existe en la conta
                b = InsertarFormaPagoEnConta(ForPago, MenError)
'--monica
'                'si hay cheque regalo comprobar q existe su forma de pago en conta
'                'sino insertarla
'                If b And ImpCheque > 0 Then b = InsertarFormaPagoEnConta(vParamAplic.ForPagoChequeRegalo, MenError)
                
                
'                CadValuesAux2 = DevuelveDesdeBDNew(conConta, "sforpa", "codforpa", "codforpa", ForPago, "N")
'                'si no existe la forma de pago en conta, insertamos la de ariges
'                If CadValuesAux2 = "" Then
'                    cadValuesAux = "tipforpa"
'                    CadValuesAux2 = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", ForPago, "N", cadValuesAux)
'                    'insertamos e sforpa de la CONTA
'                    SQL = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
'                    SQL = SQL & " VALUES(" & ForPago & ", " & DBSet(CadValuesAux2, "T") & ", " & cadValuesAux & ")"
'                    ConnConta.Execute SQL
'                End If
                
                If b Then
                    'Insertamos en la tabla scobro de la CONTA
                    Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, codbanco, codsucur, digcontr, cuentaba,text33csb,agente,departamento" ') "
                    '[Monica]22/11/2013: tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        Sql = Sql & ",iban) "
                    Else
                        Sql = Sql & ") "
                    End If
                    
                    Sql = Sql & " VALUES " & CadValues2
                    ConnConta.Execute Sql
                End If
            End If
'        End If

    End If
    
'    If b Then b = True
    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        b = False
        MenError = "Insertar en Tesoreria: " & vbCrLf & Err.Description
    End If
    InsertarEnTesoreria = b
End Function





'###### Laura 15/11/2006 Recuperar facturas ALZIRA
Public Function InsertarEnTesoreria_RecupFac(MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim Sql As String, Codmacta As String, textcsb33 As String
Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single
Dim i As Byte


    On Error GoTo EInsertarTesoreria

    b = False
    InsertarEnTesoreria_RecupFac = False
    CadValues = ""
    CadValues2 = ""

    'Obtener el Nº de Vencimientos de la forma de pago
    Sql = "SELECT numerove, primerve, restoven FROM sforpa WHERE codforpa=" & ForPago
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 Then
            'Obtener los dias de pago del cliente
            Sql = " SELECT  diapago1, diapago2, diapago3, mesnogir, diavtoat, codmacta "
            Sql = Sql & " FROM sclien "
            Sql = Sql & " WHERE codclien=" & Cliente
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            Codmacta = DBSet(Rs!Codmacta, "T")
         
            textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd,mm,yyyy") & "'"
            
            If Not Rs.EOF Then
               cadValuesAux = "('" & codTipoM & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "
               CadValuesAux2 = "('" & LetraSerie & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "
              'Primer Vencimiento
              '------------------------------------------------------------
                i = 1
                'FECHA VTO
                FecVenci = CDate(FecFactu)
                FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                If TipForPago <> 0 Then
                    FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                Else
                    FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                End If
                'Comprobar si cliente tiene mes a no girar
                FecVenci1 = FecVenci
                If CInt(DBLet(Rs!mesnogir, "N")) <> 0 Then
                    FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                End If
                
                'Comprobar si cliente tiene dia de vencimiento atrasado
                CadValues = cadValuesAux & i & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                CadValues2 = CadValuesAux2 & i & ", "
                CadValues2 = CadValues2 & Codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                
                'IMPORTE del Vencimiento
                If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFac
                Else
                    ImpVenci = Round(TotalFac / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFac Then
'                        If ImpVenci * rsVenci!numerove <> TotalFac Then
                            ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * rsVenci!numerove), 2)
'                        End If
                    End If
                End If
                CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N", "S") & ", " & DBSet(Sucursal, "N", "S") & ", " & DBSet(DigControl, "T", "S") & ", " & DBSet(CuentaBan, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                'departamento
                CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") '& ")"
                
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    CadValues2 = CadValues2 & ", " & DBSet(Iban, "T", "S") & ")"
                Else
                    CadValues2 = CadValues2 & ")"
                End If
                
                
              'Resto Vencimientos
              '--------------------------------------------------------------------
                For i = 2 To rsVenci!numerove
                  'FECHA Resto Vencimientos
                    FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
                        FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    'Comprobar si cliente tiene mes a no girar
                    FecVenci1 = FecVenci
                    If DBLet(Rs!mesnogir, "N") <> "0" Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    End If
                        
                    CadValues = CadValues & ", " & cadValuesAux & i & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    CadValues2 = CadValues2 & ", " & CadValuesAux2 & i & ", " & Codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                  'IMPORTE Resto de Vendimientos
                    ImpVenci = Round(TotalFac / rsVenci!numerove, 2)
                    CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N", "S") & ", " & DBSet(Sucursal, "N", "S") & ", " & DBSet(DigControl, "T", "S") & ", " & DBSet(CuentaBan, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N") & ", "
                    CadValues2 = CadValues2 & DBSet(Me.DirDpto, "N", "S") '& ")"
                    
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        CadValues2 = CadValues2 & ", " & DBSet(Iban, "T", "S") & ")"
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
    
    If CadValues <> "" Then
        Sql = "INSERT INTO svenci (codtipom, numfactu, fecfactu, ordefect, fecefect, impefect)"
        Sql = Sql & " VALUES " & CadValues
        Conn.Execute Sql
    End If
    
    
    'Grabar tabla scobro de la CONTABILIDAD
    '-------------------------------------------------
'    If CadValues2 <> "" Then
'        If CuentaPrev <> "" Then
'            'antes de grabar en la scobro comprobar que existe en conta.sforpa la
'            'forma de pago de la factura. Sino existe insertarla
'            'vemos si existe en la conta
'            CadValuesAux2 = DevuelveDesdeBDNew(conConta, "sforpa", "codforpa", "codforpa", ForPago, "N")
'            'si no existe la forma de pago en conta, insertamos la de ariges
'            If CadValuesAux2 = "" Then
'                cadValuesAux = "tipforpa"
'                CadValuesAux2 = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", ForPago, "N", cadValuesAux)
'                'insertamos e sforpa de la CONTA
'                SQL = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
'                SQL = SQL & " VALUES(" & ForPago & ", " & DBSet(CadValuesAux2, "T") & ", " & cadValuesAux & ")"
'                ConnConta.Execute SQL
'            End If
'
'            'Insertamos en la tabla scobro de la CONTA
'            SQL = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, codbanco, codsucur, digcontr, cuentaba,text33csb,agente,departamento) "
'            SQL = SQL & " VALUES " & CadValues2
'            ConnConta.Execute SQL
'        End If
'    End If

    b = True
    
EInsertarTesoreria:
    If Err.Number <> 0 Then b = False
    InsertarEnTesoreria_RecupFac = b
End Function
















Private Sub ReiniciarImportesFactura()
    Me.BaseImp = 0
    Me.BaseIVA1 = 0
    Me.BaseIVA2 = 0
    Me.BaseIVA3 = 0
    
    Me.BrutoFac = 0
    
    Me.ImpIVA1 = 0
    Me.ImpIVA2 = 0
    Me.ImpIVA3 = 0
    
    Me.PorceIVA1 = 0
    Me.PorceIVA2 = 0
    Me.PorceIVA3 = 0
    
    Me.TipoIVA1 = 0
    Me.TipoIVA2 = 0
    Me.TipoIVA3 = 0
    
    Me.ImpGnral = 0
    Me.ImpPPago = 0
    Me.TotalFac = 0
End Sub


Public Function CalcularDatosFactura(cadWHERE As String, NomTabla As String, NomTablaLin As String) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim Rs As ADODB.Recordset
Dim i As Integer

Dim Sql As String
Dim cadAux As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotImpIVA As Currency

Dim ImpAux As Currency
Dim ImpIVA As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA

Dim vBruto As Currency

Dim exentoIVA As Boolean
Dim conDesplaz As Boolean
    
    CalcularDatosFactura = False
    On Error GoTo ECalcular

    If codTipoM <> "FAM" Then ReiniciarImportesFactura

   '------------------------------------------------------------------------
   'En estos dos procedimientos solo entra cuando va a generar la factura y
   'codTipoMov="FAV", cuando llamamos desde pedidos, albaranes,etc, para mostrar
   'el resumen de la factura codTipoMov="ALV","PEV",... y no entra
   '-------------------------------------------------------------------------
   
   'comprobar si hay que facturar DESPLAZAMIENTOS
   CalcularDatosFactura = CalcularDesplazamiento(cadWHERE, NomTabla)
   If Not CalcularDatosFactura Then Exit Function
   
   'comprobar si hay que facturar BONIFICACIONES
    CalcularDatosFactura = CalcularBonificacion(cadWHERE, NomTablaLin)
   If Not CalcularDatosFactura Then Exit Function
   '-------------------------------------------------------------------------
   
   'comprobar si el cliente de la factura esta EXENTO de IVA
   exentoIVA = ClienteExentoIVA
   If exentoIVA Then
        cadWHERE = Replace(cadWHERE, NomTabla, NomTablaLin)
        Sql = "SELECT sum(importel) as bruto"
        Sql = Sql & " FROM " & NomTablaLin
        Sql = Sql & " WHERE " & cadWHERE
   Else
    'Agrupar el importe bruto por tipos de iva
       If codTipoM = "FAM" Then
            'para los mantenimientos obtenemos de la tabla de tipos de contrato(stipco.codartic)
            'el articulo que se va a utilizar para facturar los mantenimientos
    '        CadAux = DevuelveDesdeBDNew(conAri, "stipco", "codartic", "codtipco", TipCoMan, "T")
            If ArticFac <> "" Then 'tenemos el articulo
                Sql = "SELECT codartic,nomartic,sartic.codigiva, tiposiva.porceiva "
                Sql = Sql & " FROM sartic INNER JOIN conta" & vParamAplic.NumeroConta & ".tiposiva ON sartic.codigiva=conta" & vParamAplic.NumeroConta & ".tiposiva.codigiva"
                Sql = Sql & " WHERE codartic=" & DBSet(ArticFac, "T")
            End If
       Else
            cadWHERE = Replace(cadWHERE, NomTabla, NomTablaLin)
            Sql = "SELECT sartic.codigiva,sum(importel) as bruto"
            Sql = Sql & " FROM " & NomTablaLin & " LEFT OUTER JOIN sartic ON " & NomTablaLin & ".codartic=sartic.codartic"
            Sql = Sql & " WHERE " & cadWHERE
            Sql = Sql & " GROUP BY codigiva "
            Sql = Sql & " ORDER BY codigiva "
        End If
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    TotBruto = 0
    TotImpIVA = 0
    i = 1
    conDesplaz = False

    While Not Rs.EOF
        'Aqui vamos acumulando la suma del importe bruto de las lineas
        If codTipoM = "FAM" Then
            vBruto = BrutoFac
        Else
            vBruto = Rs!Bruto
        End If
        
        TotBruto = TotBruto + vBruto
        ImpBImIVA = vBruto
        
        If Not exentoIVA Then
            'comprobar que si hay km de DESPLAZAMIENTO se añadan al tipo de IVA correspondiente
            If TotalKm > 0 Then
                If despIVA = Rs!Codigiva Then
                    vBruto = vBruto + ImporteL
                    TotBruto = TotBruto + ImporteL
                    conDesplaz = True
                End If
            End If
        
            'comprobar si hay BONIFICACIONES que aplicar a la Factura
            If IVABonif1 = Rs!Codigiva Then
                vBruto = vBruto - TotBonif1
                TotBruto = TotBruto - TotBonif1
            ElseIf IVABonif2 = Rs!Codigiva Then
                vBruto = vBruto - TotBonif2
                TotBruto = TotBruto - TotBonif2
            ElseIf IVABonif3 = Rs!Codigiva Then
                vBruto = vBruto - TotBonif3
                TotBruto = TotBruto - TotBonif3
            End If
            
        
            'Aplicarle el dto ppago
            '---- Laura: 05/10/2006
            ImpAux = CalcularPorcentaje(vBruto, DtoPPago, 2)
'            ImpAux = CCur(CalcularDto(CStr(vBruto), CStr(DtoPPago)))
            '---- Laura: 27/09/2006
'            ImpAux = Round(ImpAux, 2)
            '----
            ImpBImIVA = vBruto - ImpAux '(bruto - DtoPP)
            
            'Aplicarle el dto grnal
            '---- Laura: 05/10/2006
            ImpAux = CalcularPorcentaje(vBruto, DtoGnral, 2)
'            ImpAux = CCur(CalcularDto(CStr(vBruto), CStr(DtoGnral)))
            '---- Laura: 27/09/2006
'            ImpAux = Round(ImpAux, 2)
            '----
            ImpBImIVA = ImpBImIVA - ImpAux '(bruto - Dtogn)
            
            'Obtener el % de IVA
            cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Rs!Codigiva), "N")
            
            'aplicar el IVA a la base imponible de ese tipo
            '---- Laura: 05/10/2006
            ImpIVA = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
'            ImpAux = CalcularDto(CStr(ImpBImIVA), cadAux)
'            ImpIVA = Round(ImpAux, 2)
            
            'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
            'los vamos acumulando
            TotImpIVA = TotImpIVA + ImpIVA
            
            
            Select Case i
                Case 1  'IVA 1
                    TipoIVA1 = Rs!Codigiva
                    
                    BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE
                    
                    PorceIVA1 = cadAux '% de IVA
                    
                    'Importe total con IVA
                    ImpIVA1 = ImpIVA
                    
                Case 2  'IVA 2
                    TipoIVA2 = Rs!Codigiva
                    
                    BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE
                    
                    PorceIVA2 = cadAux '% de IVA
                    
                    'Importe total con IVA
                    ImpIVA2 = ImpIVA
    
                Case 3  'IVA 3
                    TipoIVA3 = Rs!Codigiva
                    
                    BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE
                    
                    PorceIVA3 = cadAux '% de IVA
                    
                    'Importe total con IVA
                    ImpIVA3 = ImpIVA
            End Select
            i = i + 1
        Else 'EXENTO DE IVA
            'si hay que factura KM añadirlos al bruto
            If TotalKm > 0 Then vBruto = vBruto + ImporteL
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    'si tiene desplazamiento comprobar que se ha añadido en el IVA
    'si no calcularlo
    If TotalKm > 0 And conDesplaz = False Then
        TotBruto = TotBruto + ImporteL
        vBruto = ImporteL
        ImpBImIVA = ImporteL
        
        'Aplicarle el dto ppago
        '---- Laura: 05/10/2006
        ImpAux = CalcularPorcentaje(vBruto, DtoPPago, 2)
'        ImpAux = CCur(CalcularDto(CStr(vBruto), CStr(DtoPPago)))
        ImpBImIVA = ImpBImIVA - ImpAux '(bruto - DtoPP)
            
        'Aplicarle el dto grnal
        '---- Laura: 05/10/2006
        ImpAux = CalcularPorcentaje(vBruto, DtoGnral, 2)
'        ImpAux = CCur(CalcularDto(CStr(vBruto), CStr(DtoGnral)))
        ImpBImIVA = ImpBImIVA - ImpAux '(bruto - Dtogn)
            
        'Obtener el % de IVA
        cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(despIVA), "N")
            
        'cuota IVA: aplicar el IVA a la base imponible de ese tipo
        'Laura: 05/10/2006
        ImpAux = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
'        ImpAux = CalcularDto(CStr(ImpBImIVA), cadAux)
        ImpIVA = ImpAux
            
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotImpIVA = TotImpIVA + ImpIVA
        
        Select Case i
            Case 2 'IVA 2
                TipoIVA2 = despIVA
                
                BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA2 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA2 = ImpIVA
                
            Case 3 'IVA 3
                TipoIVA3 = Rs!Codigiva
                
                BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA3 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA3 = ImpIVA
        End Select
    End If
    
    BrutoFac = TotBruto
    
    'Aplicarle el dto ppago
    '---- Laura: 05/10/2006
    ImpPPago = CalcularPorcentaje(TotBruto, DtoPPago, 2)
'    ImpPPago = CCur(CalcularDto(CStr(TotBruto), CStr(DtoPPago)))
    '---- Laura: 27/09/2006
'    ImpPPago = Round(ImpPPago, 2)
    '----
    
    'Aplicarle el dto general
    '---- Laura: 05/10/2006
    ImpGnral = CalcularPorcentaje(TotBruto, DtoGnral, 2)
'    ImpGnral = CCur(CalcularDto(CStr(TotBruto), CStr(DtoGnral)))
    '---- Laura: 27/09/2006
'    ImpGnral = Round(ImpGnral, 2)
    '----
    
    
    'Base Imponible
    BaseImp = TotBruto - ImpPPago - ImpGnral
    
    'TOTAL de la factura
    TotalFac = BaseImp + TotImpIVA
    
    CalcularDatosFactura = True
    
ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFactura = False
    Else
        CalcularDatosFactura = True
    End If
End Function



Public Function PasarAlbaranesAFactura(TipoAlb As String, cadSQL As String, Optional ErroresAux As String, Optional EstaRecu As Boolean) As Boolean
'IN -> cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'Desde Albaranes Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETraspasoAlbFac

    PasarAlbaranesAFactura = False
    b = False
    ErroresAux = ""
    '#Laura 14/11/2006 recuperar facturas de Alzira
    'si esta recuperando facturas
    EstaRecuFac = EstaRecu
    
    Select Case TipoAlb
        Case "ALV", "AL1", "ALM": 'ALV: Albaranes venta a clientes
                    'ALM: Albaran mostrador
            codTipoM = "FAV"
        Case "ALR": 'Albaranes de reparacion en clientes
            codTipoM = "FAR"
        Case "ART": 'Albaranes de factura rectificativa
            codTipoM = "FRT"
        Case "ALS": 'Albaranes de servicios [SERVICIOS]
            codTipoM = "FAS"
    End Select
      
'--monica
'    OpeFactu = PonerTrabajadorConectado("")
'    If OpeFactu <> "" Then
'        AlmFactu = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", OpeFactu, "N")
'    Else
'        AlmFactu = "1"
'    End If
'++monica
        AlmFactu = "1"
      
    'Aqui empieza transaccion
    Conn.BeginTrans
    ConnConta.BeginTrans
   
    'Insertar la Factura (scafac,scafac1,slifac)
    MenError = "Error al Insertar en tablas de factura."
    b = InsertarFactura(cadSQL)
    
    
    'Insertar los Vencimientos de la Factura. Tabla: svenci
    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
    '#Laura: 14/11/2006 Recuperar facturas Alzira
    If b Then
        If EstaRecuFac = False Then
            MenError = "Error al actualizar en Tesoreria."
            b = InsertarEnTesoreria(MenError)
        Else
            MenError = "Error al actualizar en Tesoreria."
            b = InsertarEnTesoreria_RecupFac(MenError)
        End If
    End If
    
    'Eliminar los Albaranes incluidos en la Factura
    If b Then
        MenError = "Error al eliminar los albaranes."
        b = EliminarAlbaranes(cadSQL)
    End If
    
    'Si alguna linea de albaran tiene asociados num. serie actualizar
    'el campo numfactu de dichos num serie
    If b Then
        MenError = "Error al actualizar los Nº de serie."
        b = ActualizarNumSeries(cadSQL)
    End If
    
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        ErroresAux = Err.Description
        Err.Clear
        b = False
    End If
  
    If b Then
        Conn.CommitTrans
        ConnConta.CommitTrans
        PasarAlbaranesAFactura = True
    Else
        Conn.RollbackTrans
        ConnConta.RollbackTrans
        PasarAlbaranesAFactura = False
        ErroresAux = MenError & vbCrLf & ErroresAux & vbCrLf
    End If
    espera 0.4
End Function


Public Function PasarMtosAFactura(TipCo As String, Operador As String, mes As String, nMante As String) As Boolean
'IN -> cadSQL: cadena para seleccion de los Mantenimientos que vamos a Facturar
'Desde Mantenimientos Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETraspasoFac

    PasarMtosAFactura = False
    b = False

    codTipoM = "FAM"

    TipCoMan = TipCo
    OpeFactu = Operador
    MesFactu = mes
    numMante = nMante

    ArticFac = DevuelveDesdeBDNew(cAgro, "stipco", "codartic", "codtipco", TipCoMan, "T")
    If ArticFac = "" Then
        MenError = "El tipo de contrato no tiene articulo para facturar."
        Exit Function
    End If


    If OpeFactu <> "" Then
        AlmFactu = DevuelveDesdeBDNew(cAgro, "straba", "codalmac", "codtraba", OpeFactu, "N")
    Else
        AlmFactu = "1"
    End If


    'Aqui empieza transaccion
    Conn.BeginTrans
    ConnConta.BeginTrans


    'Insertar la Factura (scafac,scafac1,slifac)
    MenError = "Error al Insertar en tablas de factura"
    b = InsertarFactura("")

'--monica
'    If b Then
'        'Insertar en la Tabla para detalle equipos, para reimprimir
'       MenError = "Error al insertar detalle de equipos."
'       InsertarDetEquip
'    End If


    'Insertar los Vencimientos de la Factura. Tabla: svenci
    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
    If b Then
        MenError = "Error al actualizar en Tesoreria"
        b = InsertarEnTesoreria(MenError)
    End If



    'Actualizo en la tabla mantenimientos la fecha ultimio mantenimiento
    If b Then
        MenError = "Actualizar fecha ultimo mantenimiento"
        ArticFac = "UPDATE scaman SET ulmesfac = " & mes & " WHERE codclien = " & Cliente & " AND nummante= '" & numMante & "'"
        Conn.Execute ArticFac
    End If

ETraspasoFac:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Pasando Mantenimientos a Factura.", Err.Description
        b = False
    End If
    If b Then
        Conn.CommitTrans
        ConnConta.CommitTrans
        PasarMtosAFactura = True
    Else
        Conn.RollbackTrans
        ConnConta.RollbackTrans
        PasarMtosAFactura = False
    End If
    espera 0.2
End Function



Public Function PasarTicketAFactura(cadSQL As String, Optional ErroresAux As String, Optional NTicket As String, Optional NAlbTicket As String, Optional impRegalo As String) As Boolean
'IN -> cadSQL: cadena para seleccion de las ventas del ticket que vamos a Facturar
'Desde ventas Genera las Facturas de ticket correspondiente
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETrasTicFac

    OpeFactu = ErroresAux 'Aqui hemos pasado el trabajador
    NumTicket = NTicket
    NumAlbTicket = NAlbTicket '01/09/06 laura
    ImpCheque = CCur(ComprobarCero(impRegalo)) 'importe del cheque regalo (si hay)
    
    PasarTicketAFactura = False
    b = False
    ErroresAux = ""
      
      
    'por defecto se coge el almacen con el q trabaja el empleado
    If OpeFactu <> "" Then
        AlmFactu = DevuelveDesdeBDNew(cAgro, "straba", "codalmac", "codtraba", OpeFactu, "N")
    Else
        AlmFactu = "1"
    End If
    
    
    'Insertar la Factura (scafac,scafac1,slifac)
    MenError = "Error al Insertar en tablas de factura."
    b = InsertarFacturaTicket(cadSQL)
    
    
    'Insertar los Vencimientos de la Factura. Tabla: svenci
    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
    If b Then
        MenError = "Error al actualizar en Tesoreria."
        b = InsertarEnTesoreria(MenError)
    End If
    
    
    'Eliminar los Albaranes incluidos en la Factura
    If b Then
        MenError = "Error al eliminar la venta."
        b = EliminarVenta(cadSQL)
    End If
    
    
    
ETrasTicFac:
    If Err.Number <> 0 Then
        ErroresAux = Err.Description
        Err.Clear
        b = False
    End If
  
    If Not b Then ErroresAux = MenError & vbCrLf & ErroresAux & vbCrLf
    PasarTicketAFactura = b
End Function



Private Function InsertarFactura(cadSQL As String, Optional cadNumFac As String) As Boolean
'Insertamos en las tablas de factura de cliente: scafac, scafac1, slifac
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean
Dim devuelve As String

    InsertarFactura = False
    
    
    'Obtener el Contador de FACTURA
    '--------------------------------------------------------------
    Set vTipoMov = New CTiposMov
        
    If vTipoMov.leer(codTipoM) Then
        '#Laura: 14/11/2006 Recuperar facturas ALZIRA
        If EstaRecuFac = False Then
            'Comprobar si mientras tanto se incremento el contador de Facturas
            'para ello vemos si existe una Factura con ese contador y si existe la incrementamos
            Do
                NumFactu = vTipoMov.ConseguirContador(codTipoM)
                devuelve = DevuelveDesdeBDNew(cAgro, "scafac", "numfactu", "codtipom", codTipoM, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", FecFactu, "F")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (codTipoM)
                    NumFactu = vTipoMov.ConseguirContador(codTipoM)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
        End If
    Else
        Exit Function
    End If
        
    LetraSerie = vTipoMov.LetraSerie 'la guardamos para INSERT en scobro de la Conta
    
    'calcular el total de la factura antes de insertar
    b = CalcularDatosFactura(cadSQL, "scaalb", "slialb")
    If Not b Then Exit Function
    
    '### Laura: 14/11/2006 Recuperar facturas Alzira
    devuelve = ""
    If EstaRecuFac = True Then
        'comprobar q el total de la factura a insertar coincide con el total de la
        'factura en conta.cabfact si esta existe en conta
        devuelve = "SELECT COUNT(*) FROM cabfact WHERE numserie=" & DBSet(Me.LetraSerie, "T")
        devuelve = devuelve & " AND codfaccl=" & Me.NumFactu & " AND anofaccl=" & Year(FecFactu)
        If (RegistrosAListar(devuelve, cConta) > 0) Then
            devuelve = DevuelveDesdeBDNew(cConta, "cabfact", "totfaccl", "numserie", Me.LetraSerie, "T", , "codfaccl", Me.NumFactu, "N", "anofaccl", Year(Me.FecFactu), "N")
            If devuelve <> "" Then
                If CCur(devuelve) <> Me.TotalFac Then
                    If MsgBox("El total de la factura no coincide con el total en la Contabilidad" & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
                        b = False
                    End If
                End If
            End If
            devuelve = "EXISTE"
        End If
        If Not b Then Exit Function
    End If
    '#####
    
    'Cabecera Factura
    b = InsertarCabeceraFactu
    If Not b Then Exit Function
    
    'Cabeceras Albaranes de la Factura
    b = InsertarCabAlbaranesFactu(cadSQL)
    If Not b Then Exit Function
    
    'Insertar lineas de la factura
    b = InsertarLineasFactu(cadSQL)
    If Not b Then Exit Function
    
    'Comprobar si los albaranes tienen km de desplazamiento e insertar
    'lineas de factura del desplazamiento e incrementar el bruto de la factura
    'Insertar lineas de desplazamiento
    b = InsertarDesplaz
    If Not b Then Exit Function
    
    'si hay bonificaciones en la factura
    If SQLBonif <> "" Then
        b = InsertarBonificaciones
        If Not b Then Exit Function
    End If
    
   
    '#Laura: 14/11/2006 Recuperar facturas Alzira
    If EstaRecuFac = True And devuelve = "EXISTE" Then
        b = MarcarContabilizada
        If Not b Then Exit Function
    End If
    
    
    'Incrementar contador del tipo de movimiento
    '-------------------------------------------------------------
    '#Laura: 14/11/2006 Recuperar facturas Alzira
    If EstaRecuFac = False Then vTipoMov.IncrementarContador (codTipoM)
    Set vTipoMov = Nothing
      
    InsertarFactura = True
End Function




Private Function InsertarFacturaTicket(cadSQL As String) As Boolean
'Insertamos en las tablas de factura de cliente: scafac, scafac1, slifac
'Es un ticket de venta
Dim b As Boolean
Dim devuelve As String

    InsertarFacturaTicket = False
    
    
    'Calcular los campos de la factura y cuadrarla al
    'pasar las lineas sin IVA
    b = CalcularDatosFacturaTicket(cadSQL)
    If Not b Then Exit Function

    
    'Cabecera Factura
    b = InsertarCabeceraFactu
    If Not b Then Exit Function
    
    'Cabeceras Albaranes de la Factura
    b = InsertarCabAlbaranesFactu(cadSQL)
    If Not b Then Exit Function

    'Insertar lineas de la factura
    b = InsertarLineasFactu(cadSQL)
    If Not b Then Exit Function
    
       
'    'Incrementar contador del tipo de movimiento
'    '-------------------------------------------------------------
'    vTipoMov.IncrementarContador (codTipoM)
'    Set vTipoMov = Nothing
      
    InsertarFacturaTicket = True
End Function




Private Function InsertarCabeceraFactu() As Boolean
'Inserta la cabecera de la factura en la tabla: scafac
Dim Sql As String
Dim vClien As CCliente
Dim Nulo2 As String
Dim Nulo3 As String
Dim ctaDpto As Boolean
Dim cad As String, cad2 As String

    On Error GoTo EInsertar

    
    'Insertar en la tabla cabecera de la factura de compras
    Sql = "INSERT INTO scafac (codtipom,numfactu,fecfactu,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,coddirec,nomdirec,"
    Sql = Sql & "codagent,codforpa,dtoppago,dtognral,codbanco,codsucur,digcontr,cuentaba,brutofac,impdtopp,impdtogr,baseimp1,baseimp2,baseimp3, "
    Sql = Sql & "codigiv1,codigiv2,codigiv3,porciva1,porciva2,porciva3,imporiv1,imporiv2,imporiv3,totalfac,intconta, iban) "
    
    
    Sql = Sql & " VALUES (" & DBSet(codTipoM, "T") & "," & DBSet(NumFactu, "N") & "," & DBSet(FecFactu, "F") & "," & DBSet(Cliente, "N") & ","
    Sql = Sql & DBSet(NombreClien, "T") & "," & DBSet(DomicilioClien, "T", "N") & "," & DBSet(CPostal, "T", "N") & "," & DBSet(Poblacion, "T", "N") & "," & DBSet(Provincia, "T", "N") & ","
    Sql = Sql & DBSet(NIF, "T", "N") & "," & DBSet(Telefono, "T") & "," & DBSet(DirDpto, "N", "S") & "," & DBSet(NombreDirDpto, "T") & "," & DBSet(Agente, "N") & ","
    Sql = Sql & ForPago & "," & DBSet(DtoPPago, "N") & "," & DBSet(DtoGnral, "N") & ","
    
    If codTipoM = "FAV" Or codTipoM = "FRT" Or codTipoM = "FAR" Or codTipoM = "FAS" Then
        'si son facturas de VENTA o Factura Rectificativa
        'Cargar los datos del banco del cliente
        'si se trabaja con departamentos y la factura va a un departamento
        'comprobar si tiene banco y asignar el del departamento
        ctaDpto = False
'--monica
'        If vParamAplic.Departamento Then
            If Me.DirDpto <> "" Then
                'comprobar que el dpto del cliente tiene cta bancaria
                cad2 = "digcontr"
                cad = DevuelveDesdeBDNew(cAgro, "sdirec", "cuentaba", "codclien", Me.Cliente, "N", cad2, "coddirec", Me.DirDpto, "N")
                If cad <> "" And cad2 <> "" Then ctaDpto = True
            End If
'--monica
'        End If
        
        If ctaDpto Then 'cogemos la cta bancaria del departamento
            Me.CuentaBan = cad
'            cad = DevuelveDesdeBDNew(conAri, "sdirec", "digcontr", "codclien", Me.Cliente, "N", , "coddirec", Me.DirDpto, "N")
            Me.DigControl = cad2
            cad2 = "codsucur"
            cad = DevuelveDesdeBDNew(cAgro, "sdirec", "codbanco", "codclien", Me.Cliente, "N", cad2, "coddirec", Me.DirDpto, "N")
            Me.Banco = cad
            Me.Sucursal = cad2
        Else 'cuenta bancaria del cliente
            Set vClien = New CCliente
            If vClien.LeerDatos(Cliente) Then
                Banco = vClien.Banco
                Sucursal = vClien.Sucursal
                DigControl = vClien.DigControl
                CuentaBan = vClien.CuentaBan
                '[Monica]22/11/2013: Tema iban
                Iban = vClien.Iban
            Else
                Banco = ValorNulo
                Sucursal = ValorNulo
                DigControl = ValorNulo
                CuentaBan = ValorNulo
                '[Monica]22/11/2013: Tema iban
                Iban = ValorNulo
            End If
            Set vClien = Nothing
        End If
    End If
    
    Sql = Sql & DBSet(Banco, "N", "S") & "," & DBSet(Sucursal, "N", "S") & "," & DBSet(DigControl, "T", "S") & "," & DBSet(CuentaBan, "T", "S") & ","
    
    Sql = Sql & DBSet(BrutoFac, "N") & ","
    Sql = Sql & DBSet(ImpPPago, "N") & "," & DBSet(ImpGnral, "N") & ","
    Sql = Sql & DBSet(BaseIVA1, "N") & "," & DBSet(BaseIVA2, "N", "S") & "," & DBSet(BaseIVA3, "N", "S") & ","
    Nulo2 = "N"
    Nulo3 = "N"
    If BaseIVA2 = 0 Then Nulo2 = "S"
    If BaseIVA3 = 0 Then Nulo3 = "S"
    Sql = Sql & DBSet(TipoIVA1, "N") & "," & DBSet(TipoIVA2, "N", Nulo2) & "," & DBSet(TipoIVA3, "N", Nulo3) & ","
    Sql = Sql & DBSet(PorceIVA1, "N") & "," & DBSet(PorceIVA2, "N", Nulo2) & "," & DBSet(PorceIVA3, "N", Nulo3) & ","
     
    Sql = Sql & DBSet(ImpIVA1, "N") & "," & DBSet(ImpIVA2, "N", Nulo2) & "," & DBSet(ImpIVA3, "N", Nulo3) & ","
     
    Sql = Sql & DBSet(TotalFac, "N") & ",0," & DBSet(Iban, "T", "S") & ")"
    
    Conn.Execute Sql
    InsertarCabeceraFactu = True
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabeceraFactu = False
        MuestraError Err.Number, "Insertar cabecera factura.", Err.Description
    Else
        InsertarCabeceraFactu = True
    End If
End Function


Private Function InsertarCabAlbaranesFactu(cadSQL As String) As Boolean
'Insertar las cabeceras de los Albaranes de la factura en la tabla: scafac1
'NumTicket: para cuando de venta generamos Factura en numalbar guardamos el ticket que la genero
Dim Sql As String
Dim devuelve As String
Dim MantenimientoObservacion As String

    On Error GoTo EInsertarAlb

    devuelve = ""
    If codTipoM = "FTI" Then
'        cadSQL = Replace(cadSQL, "sliven", "scaven")
    Else
        If InStr(cadSQL, "scaven") > 0 Then 'viene de generar factura en ticket
'            codTipoM = "FTI"
        Else
            cadSQL = Replace(cadSQL, "slialb", "scaalb")
        End If
    End If
    
    Sql = "INSERT INTO scafac1 (codtipom,numfactu,fecfactu,codtipoa,numalbar,fechaalb,numpedcl,fecpedcl,sementre,numofert,fecofert,referenc,codenvio,codtraba,codtrab1,codtrab2,observa1,observa2,observa3,observa4,observa5,numtermi,numventa)"
    
    If (codTipoM = "FAV" And NumTicket <> "") Then
        'Facturas generadas desde una venta de tickets: en numalbar guardamos numticket
'        SQL = SQL & " VALUES('" & codTipoM & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "','FTI'," & NumTicket & ","
        '01/09/06 laura
        Sql = Sql & " VALUES('" & codTipoM & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "','ATI'," & NumTicket & ","
        
        Sql = Sql & "'" & Format(FecFactu, FormatoFecha) & "'," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de pedido a nulo
        Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de oferta a nulo
        'obtenemos el codigo de envio de cliente
        devuelve = DevuelveDesdeBDNew(cAgro, "sclien", "codenvio", "codclien", Cliente, "N")
        Sql = Sql & DBSet(devuelve, "N") & ","
        'trabajadores
        Sql = Sql & DBSet(OpeFactu, "N") & "," & ValorNulo & "," & DBSet(OpeFactu, "N") & ","
        'observaciones
        '---- Laura 10/10/06: poner en la observacion 1 el valor de scaven.observa1
        devuelve = DevuelveDesdeBDNew(cAgro, "scaven", "observa1", "numtermi", Me.NumTerminal, "N", , "numventa", Me.NumVenta, "N", "fecventa", Me.FecFactu, "F")
        Sql = Sql & DBSet(devuelve, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        
        Sql = Sql & NumTerminal & "," & NumVenta & ")"
        
    ElseIf (codTipoM <> "FAM" And codTipoM <> "FTI") And cadSQL <> "" Then 'Facturas de VEnta
        Sql = Sql & " SELECT '" & codTipoM & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu," & "codtipom as codtipoa, numalbar,fechaalb,numpedcl,fecpedcl,sementre,numofert,fecofert,referenc,codenvio,codtraba,codtrab1,codtrab2,observa01,observa02,observa03,observa04,observa05,numtermi,numventa "
        Sql = Sql & " FROM scaalb WHERE " & cadSQL
        
        
    Else 'Bonificaciones/Desplazamientos
        'Facturas de Mantenimiento
        'Facturas de Ticket del TPV
        Sql = Sql & " VALUES('" & codTipoM & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "',"
'        If (codTipoM = "FAM" Or codTipoM = "FTI") Then
        '01/09/06 laura
        If (codTipoM = "FAM") Then
            Sql = Sql & "'',0,"
        ElseIf codTipoM = "FTI" Then
            Sql = Sql & "'ATI'," & NumAlbTicket & ","
        Else
            Sql = Sql & "'',9999999,"
        End If
        Sql = Sql & "'" & Format(FecFactu, FormatoFecha) & "'," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de pedido a nulo
        Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de oferta a nulo
        'obtenemos el codigo de envio de cliente
        devuelve = DevuelveDesdeBDNew(cAgro, "sclien", "codenvio", "codclien", Cliente, "N")
        Sql = Sql & DBSet(devuelve, "N") & ","
        'trabajadores
        Sql = Sql & DBSet(OpeFactu, "N") & "," & ValorNulo & "," & DBSet(OpeFactu, "N") & ","
        'observaciones 1
        If codTipoM = "FTI" Then
            '---- Laura 19/12/06: poner en la observacion 1 el valor de scaven.observa1
            devuelve = DevuelveDesdeBDNew(cAgro, "scaven", "observa1", "numtermi", Me.NumTerminal, "N", , "numventa", Me.NumVenta, "N", "fecventa", Me.FecFactu, "F")
        Else
            'La linea viene en el campo concefac que hemos guardado en observaciones
            If codTipoM = "FAM" Then
                
                MantenimientoObservacion = DevuelveDesdeBDNew(cAgro, "scaman", "tipopago", "codclien", Me.Cliente, "N", , "nummante", numMante, "T")
                If MantenimientoObservacion = "1" Then 'Trimestral
                    MantenimientoObservacion = "TRIMESTRAL DE " & Year(FecFactu)
                ElseIf MantenimientoObservacion = "2" Then 'Semestral
                    MantenimientoObservacion = "SEMESTRAL DE " & Year(FecFactu)
                ElseIf MantenimientoObservacion = "3" Then 'Anual
                    MantenimientoObservacion = "ANUAL " & Year(FecFactu)
                Else
                    'si es mensual: ej: "JUNIO de 2005" (ponemos el mes seleccionado para facturar)
                    MantenimientoObservacion = MonthName(CLng(MesFactu)) & " de " & Year(FecFactu)
                    MantenimientoObservacion = UCase(MantenimientoObservacion)
                End If
                                                            
                If mObservacion <> "" Then
                    devuelve = MantenimientoObservacion
                    
                Else
                    devuelve = ""
                    mObservacion = MantenimientoObservacion
                End If
                
            Else
                devuelve = ""
            End If
        End If
        Sql = Sql & DBSet(devuelve, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
        Sql = Sql & "," & DBSet(NumTerminal, "N", "S") & "," & DBSet(NumVenta, "N", "S") & ")"
    End If
    
    Conn.Execute Sql

    InsertarCabAlbaranesFactu = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCabAlbaranesFactu = False
        MuestraError Err.Number, "Insertar cabecera albaranes factura.", Err.Description
    Else
        InsertarCabAlbaranesFactu = True
    End If
End Function


Private Function InsertarLineasFactu(cadSQL As String) As Boolean
'Insertar las lineas de los Albaranes de la factura en la tabla: slifac
Dim Sql As String
Dim cad As String
Dim cad2 As String
Dim cArt As CArticulo

    On Error GoTo EInsertarLin

    Sql = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre"
    If NumTicket <> "" Then Sql = Sql & ",precioiv"
    Sql = Sql & ",preciomp,preciost,preciouc)"
    
    If codTipoM = "FAM" Then 'facturas de mantenimientos
        Sql = Sql & " VALUES('" & codTipoM & "'," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','',0,1,"
            
        Sql = Sql & DBSet(AlmFactu, "N") & ","
            
        'Articulo para facturar mantenimientos de la tabla de tipos contrato: stipco.codartic
        Sql = Sql & DBSet(ArticFac, "T") & ", "
            
        'nombre del articulo para facturar,preciomp,preciost, preciouc
        If ArticFac <> "" Then
            Set cArt = New CArticulo
            If cArt.LeerDatos(ArticFac) Then
                cad = cArt.Nombre   'Aunque luego machaco cad con la nueva amplicacion: contrato + numerocontrato
            End If
'            cad2 = "preciomp"
'            cad = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", ArticFac, "T", cad2)
        End If
        
        cad = DevuelveDesdeBDNew(cAgro, "stipco", "nomtipco", "codtipco", TipCoMan, "T")
        cad = cad & "  Nº: " & numMante
        
        Sql = Sql & DBSet(cad, "T") & ","
            
        'campo Ampliacion linea segun el tipopago del mantenimientos
        '-----------------------
'        'ESTO YA LO HACE EN LA insercion de scafac1.  Con lo cual aqui pondremos lo que haya en obsraciones
'       ------------------------
'        cad = DevuelveDesdeBDNew(conAri, "scaman", "tipopago", "codclien", Me.Cliente, "N", , "nummante", numMante, "T")
'        If cad = "1" Then 'Trimestral
'            cad = "TRIMESTRAL DE " & Year(FecFactu)
'        ElseIf cad = "2" Then 'Semestral
'            cad = "SEMESTRAL DE " & Year(FecFactu)
'        ElseIf cad = "3" Then 'Anual
'            cad = "ANUAL " & Year(FecFactu)
'        Else
'            'si es mensual: ej: "JUNIO de 2005" (ponemos el mes seleccionado para facturar)
'            cad = MonthName(CLng(MesFactu)) & " de " & Year(FecFactu)
'            cad = UCase(cad)
'        End If
        cad = mObservacion
        
        Sql = Sql & DBSet(cad, "T") & ","
            
        'cantidad, precio, dtos
        Sql = Sql & "1," & DBSet(BrutoFac, "N") & ", 0,0, "
        'importe, origen precio
        Sql = Sql & DBSet(BrutoFac, "N") & ", 'M',"
        
        'guardar los precios del articulo
        Sql = Sql & DBSet(cArt.PrecioMedPon, "N") & "," & DBSet(cArt.PrecioStan, "N") & "," & DBSet(cArt.PrecioUltCom, "N") & ")"
        
        Set cArt = Nothing
        
    ElseIf (codTipoM = "FTI") Then 'factura de ticket
        'se entra en esta opcion cuando desde el TPV generamos un ticket
        Sql = Sql & " SELECT '" & codTipoM & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu, 'ATI' as codtipoa," & NumAlbTicket & " as numalbar," & "numlinea,"
        Sql = Sql & AlmFactu & " as codalmac," & "sliven.codartic,sliven.nomartic," & ValorNulo & " as ampliaci,cantidad,precioar,0 as dtoline1,0 as dtoline2, cantidad*precioar as importel,'' as origpre,precioiv "
         '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        Sql = Sql & ", sartic.preciomp,sartic.preciost,sartic.preciouc"
        '--
        Sql = Sql & " FROM sliven "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        Sql = Sql & " INNER JOIN sartic ON sliven.codartic=sartic.codartic "
        '--
        Sql = Sql & " WHERE " & Replace(cadSQL, "scaven", "sliven")
        
    ElseIf NumTicket <> "" Then 'Ticket de venta genera factura
        Sql = Sql & " SELECT '" & codTipoM & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu, 'ATI' as codtipoa," & NumTicket & " as numalbar," & "numlinea,"
        Sql = Sql & AlmFactu & " as codalmac," & "sliven.codartic,sliven.nomartic," & ValorNulo & " as ampliaci,cantidad,precioar,0 as dtoline1,0 as dtoline2, round(cantidad*precioar,2) as importel,'' as origpre,precioiv "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        Sql = Sql & ", sartic.preciomp,sartic.preciost,sartic.preciouc"
        '--
        Sql = Sql & " FROM sliven "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        Sql = Sql & " INNER JOIN sartic ON sliven.codartic=sartic.codartic "
        '--
        Sql = Sql & " WHERE " & Replace(cadSQL, "scaven", "sliven")
    
    Else 'pasamos lineas de albaranes a lineas factura
        Sql = Sql & " SELECT '" & codTipoM & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu," & "codtipom as codtipoa, numalbar,numlinea,codalmac,slialb.codartic,slialb.nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        Sql = Sql & ", sartic.preciomp,sartic.preciost,sartic.preciouc"
        '--
        Sql = Sql & " FROM slialb "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        Sql = Sql & " INNER JOIN sartic ON slialb.codartic=sartic.codartic "
        '--
        Sql = Sql & " WHERE " & Replace(cadSQL, "scaalb", "slialb")
    End If
    
    Conn.Execute Sql

    InsertarLineasFactu = True
    
EInsertarLin:
    If Err.Number <> 0 Then
        InsertarLineasFactu = False
        MuestraError Err.Number, "Insertar lineas factura.", Err.Description
    Else
        InsertarLineasFactu = True
    End If
End Function


Private Function EliminarAlbaranes(cadSQL As String) As Boolean
'Eliminamos de las tablas de Albaranes: scaalb, slialb
Dim Sql As String

    On Error GoTo EEliminar

    EliminarAlbaranes = False
    
    'ELiminar lineas albaranes
    Sql = "DELETE FROM slialb "
    Sql = Sql & " WHERE " & Replace(cadSQL, "scaalb", "slialb")
    Conn.Execute Sql
    
    espera 0.1
    
    'Eliminar Cabeceras Albaranes
    Sql = "DELETE FROM scaalb "
    Sql = Sql & " WHERE " & Replace(cadSQL, "slialb", "scaalb")
    Conn.Execute Sql
        
    EliminarAlbaranes = True

EEliminar:
    If Err.Number <> 0 Then
        EliminarAlbaranes = False
    Else
        EliminarAlbaranes = True
    End If
End Function


'--monica

Private Function ClienteExentoIVA() As Boolean
Dim vClien As CCliente
Dim b As Boolean

    Set vClien = New CCliente
    vClien.Codigo = Cliente
    b = vClien.exentoIVA
    Set vClien = Nothing
    
    ClienteExentoIVA = b
End Function



Private Function InsertarDesplaz() As Boolean
'Si los albaranes a facturar tienen KM de desplazamiento añadirlos a la factura
'Añadir cantkm * preciokm al bruto de la factura
'Insertar linea en tabla slifac con el importe del desplazamiento
Dim Sql As String
Dim cad As String
Dim b As Boolean

    On Error GoTo EDesplaz
    InsertarDesplaz = False
    
    If TotalKm > 0 Then
        'Insertar en scafac1 un albaran a 9999999, para el desplazamiento (y servira tb para bonificaciones)
        '-----------------------------------------------------------
        b = InsertarCabAlbaranesFactu("")
        If Not b Then Exit Function
            
               
        'Insertar linea de Desplazamiento en tabla: slifac (lineas factura)
        '------------------------------------------------------------------
        Sql = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
        Sql = Sql & "cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
        'codtipom,numfactu,fecfactu,codtipoa,numalbar
        Sql = Sql & " VALUES (" & DBSet(codTipoM, "T") & "," & NumFactu & "," & DBSet(FecFactu, "F") & ",'',9999999,"
        'numlinea
        Sql = Sql & "1,"
        Sql = Sql & DBSet(AlmFactu, "N") & ","
            
        'codartic,nomartic,ampliaci
        cad = DevuelveDesdeBDNew(cAgro, "sartic", "nomartic", "codartic", codArtic, "T")
        Sql = Sql & DBSet(codArtic, "T") & "," & DBSet(cad, "T") & "," & DBSet(Ampliacion, "T") & ","
        'cantidad,precioar,dtoline1,dtoline2,importel,origpre
        Sql = Sql & DBSet(TotalKm, "N") & "," & DBSet(PrecioKm, "N") & ",0,0," & DBSet(ImporteL, "N") & ",'M')"
        
        Conn.Execute Sql
        
    End If
    InsertarDesplaz = True
     
EDesplaz:
    If Err.Number <> 0 Then
        InsertarDesplaz = False
    Else
        InsertarDesplaz = True
    End If
End Function


Private Function CalcularDesplazamiento(cadWHERE As String, NomTabla As String) As Boolean
'comprobar si los albaranes a facturar tienen DESPLAZAMIENTOS
'calcula en vbles: Totalkm, Ampliacion los valores para insertar en linea factura
Dim Rs As ADODB.Recordset
Dim Sql As String
    
    CalcularDesplazamiento = False
    On Error GoTo EDesp
    
    If codTipoM = "FAV" Or codTipoM = "FAR" Then
        'Obtenemos el total de km a facturar en los albaranes seleccionados
        Sql = "SELECT fechaalb,sum(cantidkm)" ',sum(sclien.kilometr) "
        Sql = Sql & " FROM " & NomTabla '& " INNER JOIN sclien ON " & NomTabla & ".codclien=sclien.codclien "
        Sql = Sql & " WHERE " & cadWHERE '& " AND facturkm=1 "
        Sql = Sql & " GROUP BY fechaalb"
        Sql = Sql & " ORDER BY fechaalb"
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        TotalKm = 0
        Ampliacion = ""
        While Not Rs.EOF
            If Rs.Fields(1).Value <> 0 Then
                TotalKm = TotalKm + Rs.Fields(1).Value
                Ampliacion = Ampliacion & Day(Rs.Fields(0).Value) & ", "
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        If Ampliacion <> "" Then 'quitar la ultima coma
            Ampliacion = Mid(Ampliacion, 1, Len(Ampliacion) - 2)
            Ampliacion = "DIAS: " & Ampliacion
        End If
        
        If TotalKm > 0 Then
            'obtenemos el precio del km de los parametros, y el cod Articulo
            codArtic = "codartid"
            Sql = DevuelveDesdeBDNew(cAgro, "spara1", "preukmcl", "codigo", "1", "N", codArtic)
            PrecioKm = CCur(ComprobarCero(Sql))
            ImporteL = Round(TotalKm * PrecioKm, 2)
            
            'obtenemos el tipo IVA del articulo de desplazamiento
            Sql = DevuelveDesdeBDNew(cAgro, "sartic", "codigiva", "codartic", codArtic, "T")
            despIVA = CInt(ComprobarCero(Sql))
        End If
    End If
    CalcularDesplazamiento = True
    
EDesp:
    If Err.Number <> 0 Then
        CalcularDesplazamiento = False
    Else
        CalcularDesplazamiento = True
    End If
End Function



Private Function CalcularBonificacion(cadWHERE As String, NomTablaLin As String) As Boolean
'comprobar los articulos de los albaranes a facturar tienen BONIFICACION
Dim Rs As ADODB.Recordset
Dim RSaux As ADODB.Recordset
Dim Sql As String, cadIVA As String
Dim vClien As CCliente
Dim EsExcluyente As Boolean
Dim Bonif1 As Currency, Bonif2 As Currency
Dim Canti1 As Single, Canti2 As Single
Dim numlinea As Integer

    CalcularBonificacion = False
    On Error GoTo EBonif
    
    If codTipoM = "FAV" Then 'Factura de Venta
        'Obtener el cod. tarifa del cliente al que se va a facturar y
        'comprobar si para este tipo de tarifa se aplica Bonificacion
        Set vClien = New CCliente
        vClien.Codigo = Cliente
        SQLBonif = ""
        TotBonif1 = 0
        TotBonif2 = 0
        TotBonif3 = 0
        IVABonif1 = 0
        IVABonif2 = 0
        IVABonif3 = 0
            
        If vClien.Bonifica Then
            'Comprobar para los codartic de las lineas de albaran que vamos
            'a factura, para cuales hay bonificacion y la cantidad de la linea
            'cumple la cantidad para que se le aplica la bonificacion
            Sql = "SELECT slialb.codartic, sum(slialb.cantidad) as cantidad "
            Sql = Sql & "FROM " & NomTablaLin
            Sql = Sql & " INNER JOIN sbonif ON " & NomTablaLin & ".codartic=sbonif.codartic "
            Sql = Sql & " WHERE " & Replace(cadWHERE, "scaalb", NomTablaLin)
            Sql = Sql & " GROUP BY " & NomTablaLin & ".codartic "
            Sql = Sql & " order by " & NomTablaLin & ".codartic "
            
           
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            numlinea = 1
            
            While Not Rs.EOF 'para cada codartic
                'Para cada linea de slialb a la que se le aplica bonificacion
                'obtener la bonif. que se le va a aplicar

                'mirar primero si hay algun hastaca2<=cantidad
                '--------------------------------------------------
                Sql = "SELECT * from sbonif "
                Sql = Sql & " WHERE codartic=" & DBSet(Rs!codArtic, "T") '& " AND (" & RS!Cantidad & "<=sbonif.hastaca1)"
                Sql = Sql & " order by hastaca1 desc "
                Set RSaux = New ADODB.Recordset
                RSaux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                Bonif1 = 0
                Bonif2 = 0
                If Not RSaux.EOF Then
'                    If (1 <= Rs!Cantidad) And (Rs!Cantidad <= RSaux!hastaca1) Then
                    If (0 < Rs!Cantidad) And (Rs!Cantidad <= RSaux!hastaca1) Then
                        Bonif1 = RSaux!impboni1
                        Canti1 = Rs!Cantidad
                    End If
                    If (RSaux!hastaca1 < Rs!Cantidad) And (Rs!Cantidad <= RSaux!hastaca2) Then
                        Bonif1 = RSaux!impboni1
                        Canti1 = Rs!Cantidad
                        Bonif2 = RSaux!impboni2
                        Canti2 = Rs!Cantidad
                    End If
                    
                    
                    'comprobar si la bonificacion es excluyente
                    EsExcluyente = CBool(RSaux!excluyen)
                    
                    If EsExcluyente Then
                    'se aplicara una bonificacion u otra
                       If Bonif2 > 0 Then Bonif1 = 0
                    End If
                    
                    If Bonif1 > 0 Then
                        numlinea = numlinea + 1
                        SQLBonif = SQLBonif & "('" & codTipoM & "'," & NumFactu & "," & DBSet(FecFactu, "F") & ",'',9999999," & numlinea & "," & AlmFactu & "," & DBSet(RSaux!codarti1, "T") & ","
                        cadIVA = "codigiva"
                        Sql = DevuelveDesdeBDNew(cAgro, "sartic", "nomartic", "codartic", RSaux!codarti1, "T", cadIVA)
                        SQLBonif = SQLBonif & DBSet(Sql, "T") & ",''," & DBSet(Canti1, "N") & "," & DBSet(-Bonif1, "N") & ",0,0," & DBSet(0 - Round(Canti1 * Bonif1, 2), "N") & ",'M'),"
                        If IVABonif1 = 0 Or IVABonif1 = CByte(cadIVA) Then
                            IVABonif1 = CByte(cadIVA)
                            TotBonif1 = TotBonif1 + Round(Canti1 * Bonif1, 2)
                        ElseIf IVABonif2 = 0 Or IVABonif2 = CByte(cadIVA) Then
                            IVABonif2 = CByte(cadIVA)
                            TotBonif2 = TotBonif2 + Round((Canti1 * Bonif1), 2)
                        ElseIf IVABonif3 = 0 Or IVABonif3 = CByte(cadIVA) Then
                            IVABonif3 = CByte(cadIVA)
                            TotBonif3 = TotBonif3 + Round(Canti1 * Bonif1, 2)
                        End If
                    End If
                    
                    If Bonif2 > 0 Then
'                        SQL = "(codtipom='" & codTipoM & "' and numfactu=" & NumFactu & " and fecfactu=" & DBSet(FecFactu, "F") & " and codtipoa='' and numalbar=0)"
                        numlinea = numlinea + 1
                        SQLBonif = SQLBonif & "('" & codTipoM & "'," & NumFactu & "," & DBSet(FecFactu, "F") & ",'',9999999," & numlinea & "," & AlmFactu & "," & DBSet(RSaux!codArti2, "T") & ","
                        Sql = DevuelveDesdeBDNew(cAgro, "sartic", "nomartic", "codartic", RSaux!codArti2, "T")
                        SQLBonif = SQLBonif & DBSet(Sql, "T") & ",''," & DBSet(Canti2, "N") & "," & DBSet(-Bonif2, "N") & ",0,0," & DBSet(0 - Round(Canti2 * Bonif2, 2), "N") & ",'M'),"
                        If IVABonif1 = 0 Or IVABonif1 = CByte(cadIVA) Then
                            IVABonif1 = CByte(cadIVA)
                            TotBonif1 = TotBonif1 + Round((Canti2 * Bonif2), 2)
                        ElseIf IVABonif2 = 0 Or IVABonif2 = CByte(cadIVA) Then
                            IVABonif2 = CByte(cadIVA)
                            TotBonif2 = TotBonif2 + Round((Canti2 * Bonif2), 2)
                        ElseIf IVABonif3 = 0 Or IVABonif3 = CByte(cadIVA) Then
                            IVABonif3 = CByte(cadIVA)
                            TotBonif3 = TotBonif3 + Round((Canti2 * Bonif2), 2)
                        End If
                    End If
'                    RSaux.MoveNext
                End If
                RSaux.Close
                Set RSaux = Nothing
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            If SQLBonif <> "" Then 'quitamos ultima coma
                SQLBonif = Mid(SQLBonif, 1, Len(SQLBonif) - 1)
            End If
        End If
    End If
    Set vClien = Nothing
    CalcularBonificacion = True
    
    
EBonif:
    If Err.Number <> 0 Then
        CalcularBonificacion = False
    Else
        CalcularBonificacion = True
    End If
End Function



Private Function InsertarBonificaciones() As Boolean
'Si los albaranes a facturar tienen lineas de articulos con bonificacion
'añadirlos a la factura
'Insertar lineas en tabla slifac con el importe de las bonificaciones
Dim Sql As String
Dim b As Boolean

    On Error GoTo EBonifica

    InsertarBonificaciones = False
    
    If SQLBonif > "" Then
        'Insertar en scafac1 un albaran a 0, para la bonificacion
        'si no se inserto antes para los desplazamientos
        '-----------------------------------------------------------
        Sql = "(codtipom='" & codTipoM & "' and numfactu=" & NumFactu & " and fecfactu=" & DBSet(FecFactu, "F") & " and codtipoa='' and numalbar=9999999)"
        Sql = "Select count(*) from scafac1 where " & Sql
        If RegistrosAListar(Sql) = 0 Then
            b = InsertarCabAlbaranesFactu("")
            If Not b Then Exit Function
        End If
               
        'Insertar lineas de Desplazamiento en tabla: slifac (lineas factura)
        '------------------------------------------------------------------
        Sql = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
        Sql = Sql & "cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
        Sql = Sql & " VALUES " & SQLBonif
        Conn.Execute Sql
        
    End If
    InsertarBonificaciones = True
     
EBonifica:
    If Err.Number <> 0 Then
        InsertarBonificaciones = False
    Else
        InsertarBonificaciones = True
    End If
End Function




Private Function ActualizarNumSeries(cadSQL As String) As Boolean
'Actualiza el campo "numfactu" de la tabla de num. serie: sserie
'Para ponerle la factura asociada
Dim Sql As String

    On Error GoTo EActualizar

    Sql = "UPDATE sserie SET numfactu= " & NumFactu
    Sql = Sql & " WHERE " & Replace(cadSQL, "scaalb", "sserie")
    Conn.Execute Sql

    ActualizarNumSeries = True
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarNumSeries = False
    Else
        ActualizarNumSeries = True
    End If
End Function



Private Function CalcularDatosFacturaTicket(cadSQL As String) As Boolean
Dim b As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim i As Byte
Dim ImpTotal As Currency
Dim ImpDif As Currency

    On Error GoTo ETicket
    
'    'Bruto de la factura sera la suma de cantidad*precioar de las linea
'    'importe sin iva
'    SQL = "SELECT sum(cantidad*precioar) FROM sliven WHERE "
'    SQL = SQL & Replace(cadSQL, "scaven", "sliven")
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not RS.EOF Then
'        BrutoFac = Round(RS.Fields(0).Value, 2)
'        BaseImp = BrutoFac
'    End If
'    RS.Close
'    Set RS = Nothing
    
    
    
    'desglosar las bases imponibles por tipo de IVA
'    SQL = "SELECT codigiva,sum(cantidad*precioar) as baseimp from sliven "
    '01/09/06
    Sql = "SELECT codigiva,sum(importel) as totalimp from sliven "
    
    Sql = Sql & " WHERE " & Replace(cadSQL, "scaven", "sliven")
    Sql = Sql & " group by codigiva"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 1
    While Not Rs.EOF And i <= 3
        Sql = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Rs!Codigiva), "N")
        Select Case i
            Case 1
                TipoIVA1 = Rs.Fields(0).Value
                PorceIVA1 = CCur(Sql)
                BaseIVA1 = Round2(Rs.Fields(1).Value / (1 + (PorceIVA1 / 100)), 2)
'                SQL = CalcularDto(CStr(BaseIVA1), CStr(PorceIVA1))
'                ImpIVA1 = Round(CCur(SQL), 2)
                ImpIVA1 = Rs.Fields(1).Value - BaseIVA1
            Case 2
                TipoIVA2 = Rs.Fields(0).Value
                PorceIVA2 = CCur(Sql)
'                BaseIVA2 = RS.Fields(1).Value
'                SQL = CalcularDto(CStr(BaseIVA2), CStr(PorceIVA2))
'                ImpIVA2 = CCur(SQL)
                BaseIVA2 = Round2(Rs.Fields(1).Value / (1 + (PorceIVA2 / 100)), 2)
                ImpIVA2 = Rs.Fields(1).Value - BaseIVA2
            Case 3
                TipoIVA3 = Rs.Fields(0).Value
                PorceIVA3 = CCur(Sql)
'                BaseIVA3 = RS.Fields(1).Value
'                SQL = CalcularDto(CStr(BaseIVA3), CStr(PorceIVA3))
'                ImpIVA3 = CCur(SQL)
                BaseIVA3 = Round2(Rs.Fields(1).Value / (1 + (PorceIVA3 / 100)), 2)
                ImpIVA3 = Rs.Fields(1).Value - BaseIVA3
        End Select
        Rs.MoveNext
        i = i + 1
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    
    
    'comprobar que suma bases imp + imporiva = total factura
    ImpTotal = (BaseIVA1 + ImpIVA1) + (BaseIVA2 + ImpIVA2) + (BaseIVA3 + ImpIVA3)
    BrutoFac = BaseIVA1 + BaseIVA2 + BaseIVA3
    BaseImp = BrutoFac
    
    If Not (ImpTotal = TotalFac) Then
        ImpDif = TotalFac - ImpTotal
        ImpIVA1 = ImpIVA1 + ImpDif
'        If BaseIVA2 = 0 And BaseIVA3 = 0 Then ImpIVA1 = ImpIVA1 + ImpDif
'        ImpTotal = (BaseIVA1 + ImpIVA1) + (BaseIVA2 + ImpIVA2) + (BaseIVA3 + ImpIVA3)
    End If
    
    b = True
    
    
ETicket:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Calculando datos de ticket a factura.", Err.Description
        b = False
    End If
    CalcularDatosFacturaTicket = b
End Function


Private Function InsertarFormaPagoEnConta(nForPa As String, cadErr As String) As Boolean
Dim cadAux As String
Dim cadAux2 As String
Dim Sql As String

    On Error GoTo ErrInsForpa
    InsertarFormaPagoEnConta = False
    
    'antes de grabar en la scobro comprobar que existe en conta.sforpa la
    'forma de pago de la factura. Sino existe insertarla
    
    'vemos si existe en la conta
    cadAux = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", nForPa, "N")
    'si no existe la forma de pago en conta, insertamos la de ariges
    If cadAux = "" Then
        cadAux2 = "tipforpa"
        cadAux = DevuelveDesdeBDNew(cAgro, "sforpa", "nomforpa", "codforpa", nForPa, "N", cadAux2)
        If cadAux <> "" Then
            'insertamos e sforpa de la CONTA
            Sql = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
            Sql = Sql & " VALUES(" & nForPa & ", " & DBSet(cadAux, "T") & ", " & cadAux2 & ")"
            ConnConta.Execute Sql
            InsertarFormaPagoEnConta = True
        Else
            InsertarFormaPagoEnConta = False
        End If
    Else
        InsertarFormaPagoEnConta = True
    End If
    
    
    Exit Function
    
ErrInsForpa:
    InsertarFormaPagoEnConta = False
    cadErr = "Insertar forma de pago en Contablilidad: " & vbCrLf & Err.Description
End Function



'#### Laura 15/11/2006 Recuperar facturas ALZIRA
Public Function MarcarContabilizada(Optional cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    MarcarContabilizada = False
    
    Sql = "UPDATE scafac SET intconta=1 "
    Sql = Sql & " WHERE codtipom=" & DBSet(Me.codTipoM, "T") & " AND numfactu=" & Me.NumFactu & " AND fecfactu=" & DBSet(Me.FecFactu, "F")
    Conn.Execute Sql
    
    MarcarContabilizada = True
    Exit Function
    
EActualizar:
    MarcarContabilizada = False
    cadErr = Err.Description
End Function




