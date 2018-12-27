VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFacturaVta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'VENTAS
'FACTURA: se corresponde con la tabla ariges.scafac (cabecera de facturas)
'-------------------------------------------------------------------------

Private WithEvents frmMens As frmMensajes ' devolvemos las facturas a cuenta que vamos a descontar
Attribute frmMens.VB_VarHelpID = -1

Dim FacturasADescontar As String

'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodTipoM As String 'Tipo de Movimiento="FAV", "FAR" (Factura venta o reparacion)
Private mNumFactu As Long 'Nº Factura
Private mFecFactu As String 'FEcha Factura
Private mCodClien As String 'Cod. cliente
Private mCodSocio As String 'Cod. socio

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
'[Monica]22/11/2013: Tema iban
Private mIban As String 'codigo iban


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


Private mTipoMoneda As Byte
Private mCambioDivisa As Currency




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

'[Monica]22/11/2013: iban de la factura
'**** Iban de la Factura
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



'**** Socio de la Factura

Public Property Let Socio(ByVal vData As String)
     mCodSocio = vData
End Property

Public Property Get Socio() As String
     Socio = mCodSocio
End Property



'**** Tipo de moneda

Public Property Let TipoMoneda(ByVal vData As String)
     mTipoMoneda = vData
End Property

Public Property Get TipoMoneda() As String
     TipoMoneda = mTipoMoneda
End Property


'**** Cambio Divisa

Public Property Let CambioDivisa(ByVal vData As Currency)
     mCambioDivisa = vData
End Property

Public Property Get CambioDivisa() As Currency
     CambioDivisa = mCambioDivisa
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
    Sql = "SELECT codtipom,numfactu,fecfactu,codclien,codforpa,codbanco,codsucur,digcontr,cuentaba,brutofac, totalfac, iban FROM scafac "
    Sql = Sql & " WHERE codtipom='" & vCodtipom & "'"
    Sql = Sql & " AND numfactu=" & vNumfactu
    Sql = Sql & " AND fecfactu='" & Format(vFecfactu, FormatoFecha) & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
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
        TipForPago = DevuelveDesdeBDNew(cAgro, "forpago", "tipforpa", "codforpa", ForPago, "N")
        mCodbanco = CStr(DBLet(Rs!codbanco, "N"))
        mCodsucur = CStr(DBLet(Rs!codsucur, "N"))
        mDigContr = DBLet(Rs!digcontr, "T")
        mCuentaba = DBLet(Rs!cuentaba, "T")
        '[Monica]22/11/2013: tema iban
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
            If vParamAplic.ContabilidadNueva Then
                Sql = "SELECT COUNT(*) FROM cobros WHERE numserie='" & LetraSerie & "' and numfactu=" & NumFactu
                Sql = Sql & " AND fecfactu=" & DBSet(FecFactu, "F")
            Else
                Sql = "SELECT COUNT(*) FROM scobro WHERE numserie='" & LetraSerie & "' and codfaccl=" & NumFactu
                Sql = Sql & " AND fecfaccl=" & DBSet(FecFactu, "F")
            End If
            If RegistrosAListar(Sql, cConta) > 0 Then
                If vParamAplic.ContabilidadNueva Then
                    Sql = DevuelveDesdeBDNew(cConta, "cobros", "ctabanc1", "numserie", mLetraSer, "T", , "numfactu", vNumfactu, "N", "fecfactu", vFecfactu, "F")
                Else
                    Sql = DevuelveDesdeBDNew(cConta, "scobro", "ctabanc1", "numserie", mLetraSer, "T", , "codfaccl", vNumfactu, "N", "fecfaccl", vFecfactu, "F")
                End If
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
    Sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & ForPago
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 Then
            'Obtener los dias de pago del cliente
            Sql = " SELECT  diapago1, diapago2, diapago3, mesnogir, diavtoat, codmacta "
            Sql = Sql & " FROM sclien "
            Sql = Sql & " WHERE codclien=" & Cliente
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
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
        conn.Execute Sql
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
    rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 Then
            'Obtener los dias de pago del cliente
            Sql = " SELECT  diapago1, diapago2, diapago3, mesnogir, diavtoat, codmacta "
            Sql = Sql & " FROM sclien "
            Sql = Sql & " WHERE codclien=" & Cliente
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
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
                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                'departamento
                CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ")"
                
                
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
                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N") & ", "
                    CadValues2 = CadValues2 & DBSet(Me.DirDpto, "N", "S") & ")"
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
        conn.Execute Sql
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


Private Function CalcularDatosFacturaVenta(vCliente As CCliente) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim Rs As ADODB.Recordset
Dim i As Integer

Dim Sql As String
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

Dim cadwhere As String
Dim cadwhere1 As String
    
Dim Nulo2 As String
Dim Nulo3 As String

    
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
    cadwhere = "facturas.codtipom = " & DBSet(codTipoM, "T") & " and facturas.numfactu = " & DBSet(NumFactu, "N") & " and facturas.fecfactu = " & DBSet(FecFactu, "F")
    cadwhere1 = Replace(cadwhere, "facturas", "facturas_variedad")
    
    Sql = "SELECT facturas_variedad.codigiva, sum(imporbru) as bruto, sum(impornet) as neto"
    Sql = Sql & " FROM facturas_variedad "
    Sql = Sql & " WHERE " & cadwhere1
    Sql = Sql & " GROUP BY 1 "
    Sql = Sql & " UNION "
    cadwhere1 = Replace(cadwhere, "facturas", "facturas_envases")
    Sql = Sql & "SELECT facturas_envases.codigiva, sum(importel) as bruto, sum(importel) as neto"
    Sql = Sql & " FROM facturas_envases "
    Sql = Sql & " WHERE " & cadwhere1
    Sql = Sql & " GROUP BY 1 "
'[Monica]25/10/2011: añadido que descuente las facturas a cuenta
'    sql = sql & " ORDER BY 1 "
    Sql = Sql & " UNION "
    cadwhere1 = Replace(cadwhere, "facturas", "facturas_acuenta")
    Sql = Sql & "SELECT facturas.codiiva1 as codigiva, sum(brutofac * (-1)) as bruto, sum(brutofac * (-1)) as neto"
    Sql = Sql & " FROM facturas_acuenta, facturas "
    Sql = Sql & " WHERE " & cadwhere1
    Sql = Sql & " and facturas.codtipom = facturas_acuenta.codtipomcta "
    Sql = Sql & " and facturas.numfactu = facturas_acuenta.numfactucta "
    Sql = Sql & " and facturas.fecfactu = facturas_acuenta.fecfactucta "
    Sql = Sql & " GROUP BY 1 "
    Sql = Sql & " ORDER BY 1 "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

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

            If vCliente.TipoIva = 2 Then
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
    
    '[Monica]07/04/2015: si no encuentra el porcentaje de iva damos un mensaje de error y no se crea la factura
    If cadAux = "" Then
        MsgBox "No existe el iva en la contabilidad. Revise.", vbExclamation
        Exit Function
    End If

    'aplicar el IVA a la base imponible de ese tipo
    impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
    
    'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
    'los vamos acumulando
    TotImpIVA = TotImpIVA + impiva
    
    If CInt(vCliente.TipoIva) = 2 Then
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

    Nulo2 = "N"
    Nulo3 = "N"
    If DBSet(TipoIVA2, "N", "S") = ValorNulo Then Nulo2 = "S"
    If DBSet(TipoIVA3, "N", "S") = ValorNulo Then Nulo3 = "S"
    
    'ACTUALIZAMOS LA FACTURA (tabla facturas)
    Sql = "update facturas "
    Sql = Sql & "set baseimp1 = " & DBSet(BaseIVA1, "N")
    Sql = Sql & ",baseimp2 = " & DBSet(BaseIVA2, "N", Nulo2)
    Sql = Sql & ",baseimp3 = " & DBSet(BaseIVA3, "N", Nulo3)
    Sql = Sql & ",impoiva1 = " & DBSet(ImpIVA1, "N")
    Sql = Sql & ",impoiva2 = " & DBSet(ImpIVA2, "N", Nulo2)
    Sql = Sql & ",impoiva3 = " & DBSet(ImpIVA3, "N", Nulo3)
    Sql = Sql & ",imporec1 = " & DBSet(ImpREC1, "N")
    Sql = Sql & ",imporec2 = " & DBSet(ImpREC2, "N", Nulo2)
    Sql = Sql & ",imporec3 = " & DBSet(ImpREC3, "N", Nulo3)
    Sql = Sql & ",codiiva1 = " & DBSet(TipoIVA1, "N")
    Sql = Sql & ",codiiva2 = " & DBSet(TipoIVA2, "N", Nulo2)
    Sql = Sql & ",codiiva3 = " & DBSet(TipoIVA3, "N", Nulo3)
    Sql = Sql & ",porciva1 = " & DBSet(PorceIVA1, "N")
    Sql = Sql & ",porciva2 = " & DBSet(PorceIVA2, "N", Nulo2)
    Sql = Sql & ",porciva3 = " & DBSet(PorceIVA3, "N", Nulo3)
    Sql = Sql & ",porcrec1 = " & DBSet(PorceREC1, "N")
    Sql = Sql & ",porcrec2 = " & DBSet(PorceREC2, "N", Nulo2)
    Sql = Sql & ",porcrec3 = " & DBSet(PorceREC3, "N", Nulo3)
    Sql = Sql & ",brutofac = " & DBSet(TotBruto, "N")
    Sql = Sql & ",impordto = " & DBSet(Round2(TotBruto - TotNeto, 2), "N")
    Sql = Sql & ",totalfac = " & DBSet(TotalFac, "N")
    Sql = Sql & " where " & cadwhere
    
    conn.Execute Sql

    CalcularDatosFacturaVenta = True
    Exit Function
    
ECalcular:
    '[Monica]07/04/2015: no dabamos el mensaje de error, he quitado lo de abajo y he puesto el exit function
    MuestraError Err.Number, "Calcular Datos Factura Venta", Err.Description
    
'    If Err.Number <> 0 Then
'        CalcularDatosFacturaVenta = False
'    Else
'        CalcularDatosFacturaVenta = True
'    End If
End Function



Public Function PasarAlbaranAFactura(cadSQL As String, FechaFact As String, Optional ErroresAux As String, Optional cTipoM As String, Optional numFac As String, Optional fecFac As String, Optional FPago As String, Optional TipoMoneda As String, Optional CambioDivisa As String) As Boolean
'IN -> cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'Desde Albaranes Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETraspasoAlbFac

    PasarAlbaranAFactura = False
    b = False
    ErroresAux = ""
      
    'Aqui empieza transaccion
    conn.BeginTrans
    ConnConta.BeginTrans
   
    FecFactu = FechaFact
   
    'Insertar la Factura (scafac,scafac1,slifac)
    MenError = "Error al Insertar en tablas de factura."
    b = InsertarFactura(cadSQL)
    
    cTipoM = codTipoM
    numFac = CStr(NumFactu)
    fecFac = CStr(FecFactu)
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        ErroresAux = Err.Description
        Err.Clear
        b = False
    End If
  
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        PasarAlbaranAFactura = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        PasarAlbaranAFactura = False
        ErroresAux = MenError & vbCrLf & ErroresAux & vbCrLf
    End If
    espera 0.4
End Function


Private Function InsertarFactura(cadSQL As String) As Boolean
'Insertamos en las tablas de factura de cliente: scafac, scafac1, slifac
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean
Dim devuelve As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim vCliente As CCliente
Dim TipoMovimiento As String
Dim NumFactuNue As String
Dim NumFactuAnt As String

    InsertarFactura = False
    
    'cogemos los datos del albaran
    Sql = "select codclien, fechaalb, observac from albaran where " & cadSQL
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Set vCliente = New CCliente
    
    'si se ha modificado el cliente volver a cargar los datos
    If vCliente.Existe(Rs.Fields(0).Value) Then
        If vCliente.LeerDatos(Rs.Fields(0).Value) Then
             Cliente = vCliente.Codigo
            
            '[Monica]29/12/2017: para el caso de IMG miro si el albaran tiene alguna variedad BIO, si  Si cogerá distinto contador y tipo de movimiento
            If vParamAplic.Cooperativa = 15 And EsAlbaranBIO(cadSQL) Then
                TipoMovimiento = "FBI"
                
                Set vTipoMov = New CTiposMov
                If vTipoMov.Leer(TipoMovimiento) Then
                    NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                    NumFactuAnt = NumFactu
                    ' si existe la factura incrementamos el contador
                    Do
                        devuelve = DevuelveDesdeBDNew(cAgro, "facturas", "numfactu", "codtipom", TipoMovimiento, "T", , "numfactu", NumFactu, "N", "fecfactu", FecFactu, "F")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (TipoMovimiento)
                            NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                Else
                    Exit Function
                End If
            
            Else
                If vCliente.TipoFactu = 0 Then
                    ' 0 = automatica 1 = manual
                    TipoMovimiento = vCliente.tipoMov
                    Set vTipoMov = New CTiposMov
                    If vTipoMov.Leer(TipoMovimiento) Then
                        NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                        NumFactuAnt = NumFactu
                        ' si existe la factura incrementamos el contador
                        Do
                            devuelve = DevuelveDesdeBDNew(cAgro, "facturas", "numfactu", "codtipom", TipoMovimiento, "T", , "numfactu", NumFactu, "N", "fecfactu", FecFactu, "F")
                            If devuelve <> "" Then
                                'Ya existe el contador incrementarlo
                                Existe = True
                                vTipoMov.IncrementarContador (TipoMovimiento)
                                NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                            Else
                                Existe = False
                            End If
                        Loop Until Not Existe
                    Else
                        Exit Function
                    End If
                Else
                    ' pedimos el tipo de movimiento y el numero de factura
    '++monica:040608: pedimos unicamente el tipo de movimiento
                    TipoMovimiento = UCase(InputBox("Tipo Factura:", "Tipo Factura:"))
    '--monica:27/11/2009
    '                If MsgBox("¿Datos correctos para continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
    '                    Exit Function
    '                End If
                    Set vTipoMov = New CTiposMov
                    If vTipoMov.Leer(TipoMovimiento) Then
                        NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                        NumFactuAnt = NumFactu
                        NumFactuNue = InputBox("Nro.Factura:", "Nro.Factura:", NumFactu)
                        ' si existe la factura incrementamos el contador
                        If NumFactu = NumFactuNue Then
                            Do
                                devuelve = DevuelveDesdeBDNew(cAgro, "facturas", "numfactu", "codtipom", TipoMovimiento, "T", , "numfactu", NumFactu, "N", "fecfactu", FecFactu, "F")
                                If devuelve <> "" Then
                                    'Ya existe el contador incrementarlo
                                    Existe = True
                                    vTipoMov.IncrementarContador (TipoMovimiento)
                                    NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                                Else
                                    Existe = False
                                End If
                            Loop Until Not Existe
                        Else
                            devuelve = DevuelveDesdeBDNew(cAgro, "facturas", "numfactu", "codtipom", TipoMovimiento, "T", , "numfactu", NumFactuNue, "N", "fecfactu", FecFactu, "F")
                            If devuelve <> "" Then
                                MsgBox "Factura existente. Revise.", vbExclamation
                                Exit Function
                            End If
                            NumFactu = NumFactuNue
                        End If
                        If MsgBox("¿Datos correctos para continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                    
                End If
                
            End If
            
            codTipoM = TipoMovimiento
            
            'Cabecera Factura
            b = InsertarCabeceraFactu(vCliente, DBLet(Rs.Fields(2).Value, "T"))
            If Not b Then Exit Function
             
            'Insertar lineas de variedades de la factura
            b = InsertarLineasFactu(cadSQL, vCliente)
            If Not b Then Exit Function
                         
            '[Monica]04/04/2012: faltaba las lineas de calibre
            'Insertamos las lineas de calibre de cada una de las lineas de factura
            b = InsertarLineasCalibreFactu(cadSQL)
            If Not b Then Exit Function
                         
            b = RecalcularDtos(vCliente)
            If Not b Then Exit Function
            
            '[Monica]25/10/2011
            b = DescontarFacturasACuenta(codTipoM, NumFactu, FecFactu, Rs.Fields(0).Value)
            If Not b Then Exit Function
            
            'calcular el total de la factura
            b = CalcularDatosFacturaVenta(vCliente)
            If Not b Then Exit Function
            
            'Incrementar contador del tipo de movimiento
            '-------------------------------------------------------------
'--monica: en cualquier caso se incrementa contador
'            If vCliente.TipoFactu = 0 Then
            If NumFactuAnt = NumFactu Then
                vTipoMov.IncrementarContador (codTipoM)
                Set vTipoMov = Nothing
            End If
            
            InsertarFactura = True
            
        End If
    End If
    Set vCliente = Nothing
    
End Function


Private Function InsertarCabeceraFactu(ByRef vCliente As CCliente, Observac As String) As Boolean
'Inserta la cabecera de la factura en la tabla: scafac
Dim Sql As String
Dim vClien As CCliente
Dim Nulo2 As String
Dim Nulo3 As String
Dim ctaDpto As Boolean
Dim Cad As String, cad2 As String

    On Error GoTo EInsertar


    'Insertar en la tabla cabecera de la factura de ventas
    Sql = "INSERT INTO facturas (codtipom,numfactu,fecfactu,codclien,codforpa,tipoivac,dtocom1,"
    Sql = Sql & "dtocom2,impdtoc,observac,baseimp1,baseimp2,baseimp3,impoiva1,impoiva2,impoiva3,"
    Sql = Sql & "imporec1,imporec2,imporec3,codiiva1,codiiva2,codiiva3,porciva1,porciva2,porciva3,"
    Sql = Sql & "porcrec1,porcrec2,porcrec3,brutofac,impordto,totalfac,intconta,pasaridoc, "
    '[Monica]17/12/2018: añadimos el tipo de moneda y el cambio moneda
    Sql = Sql & "codmoneda, cambiodivisa)"
    
    Sql = Sql & " VALUES (" & DBSet(codTipoM, "T") & "," & DBSet(NumFactu, "N") & "," & DBSet(FecFactu, "F") & "," & DBSet(vCliente.Codigo, "N") & ","
    If mCodForpa <> "" Then
        Sql = Sql & DBSet(mCodForpa, "N")
    Else
        Sql = Sql & DBSet(vCliente.ForPago, "N")
    End If
    
    Sql = Sql & "," & DBSet(vCliente.TipoIva, "N") & "," & DBSet(vCliente.Dto1, "N") & "," & DBSet(vCliente.Dto2, "N") & ","
    Sql = Sql & ValorNulo & "," & DBSet(Observac, "T", "S") & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & "0,0,"
    
    '[Monica]17/12/2018: tipo de moneda y cambio moneda
    Sql = Sql & DBSet(mTipoMoneda, "N") & "," & DBSet(mCambioDivisa, "N")
    Sql = Sql & ")"
    
    conn.Execute Sql
    
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
    
    conn.Execute Sql

    InsertarCabAlbaranesFactu = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCabAlbaranesFactu = False
        MuestraError Err.Number, "Insertar cabecera albaranes factura.", Err.Description
    Else
        InsertarCabAlbaranesFactu = True
    End If
End Function


Private Function InsertarLineasFactu(vSQL As String, ByRef vCliente As CCliente) As Boolean
'Insertar las lineas de los Albaranes de la factura en la tabla: slifac
Dim Sql As String
Dim Sql2 As String
Dim Cad As String
Dim cad2 As String
Dim vWhere2 As String
Dim cArt As CArticulo
Dim Rs As ADODB.Recordset
Dim NumF As Long
Dim TipoIva As Integer
Dim vSql1 As String
Dim vPrecio As Currency

    On Error GoTo EInsertarLin

    Sql = "INSERT INTO facturas_variedad (codtipom,numfactu,fecfactu,numlinea,numalbar,"
    Sql = Sql & "numlinealbar,cantreal,cantfact,precibru,precinet,dtocom1,dtocom2,imporbru,"
    Sql = Sql & "impornet,codigiva,unidades) values "
    
    vWhere2 = "codtipom = " & DBSet(codTipoM, "T") & " and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
    
    '[Monica]11/10/2018: si el cliente tiene un tipo de movimiento FXX (B) iva exento
    If vCliente.tipoMov = TipoFactB Then
        TipoIva = vParamAplic.CodIvaExento
    Else
        Select Case vCliente.TipoIva
            Case 0  'normal
    '            TipoIva = vParamAplic.CodIvaNormal
    '++monica:27/05/08 el tipo de iva lo sacaré de la variedad
                TipoIva = 0
            Case 1  'exento
                TipoIva = vParamAplic.CodIvaExento
            Case 2  'recargo de equivalencia
                TipoIva = vParamAplic.CodIvaRecargo
        End Select
    End If
    
    vSql1 = Replace(vSQL, "albaran", "albaran_variedad")
    
    Sql2 = "select albaran_variedad.numalbar, albaran_variedad.numlinea, albaran_variedad.codforfait,"
    Sql2 = Sql2 & " albaran_variedad.totpalet, albaran_variedad.numcajas, albaran_variedad.pesobrut,"
    Sql2 = Sql2 & " albaran_variedad.pesoneto, albaran_variedad.preciopro, albaran_variedad.preciodef,"
    Sql2 = Sql2 & " albaran_variedad.impcomis, forfaits.kiloscaj, albaran_variedad.unidades, "
    Sql2 = Sql2 & " albaran_variedad.codvarie, forfaits.facturar "
    Sql2 = Sql2 & " from albaran_variedad, forfaits where " & vSql1
    Sql2 = Sql2 & " and albaran_variedad.codforfait = forfaits.codforfait "
    Sql2 = Sql2 & " order by albaran_variedad.numalbar, albaran_variedad.numlinea"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Cad = "(" & DBSet(codTipoM, "T") & "," & DBSet(NumFactu, "N") & "," & DBSet(FecFactu, "F") & ","
    cad2 = ""
    
    NumF = 1
    
    While Not Rs.EOF
        '[Monica]23/04/2013: cogemos el definitivo si lo tiene, si no el provisional que si no hay es cero
        If DBLet(Rs.Fields(8).Value, "N") <> 0 Then
            vPrecio = DBLet(Rs.Fields(8).Value, "N")
        Else
            vPrecio = DBLet(Rs.Fields(7).Value, "N")
        End If
        
    
        If vCliente.TipoIva = 0 Then
            TipoIva = DevuelveDesdeBDNew(cAgro, "variedades", "codigiva", "codvarie", Rs.Fields(12).Value, "N")
        End If
        cad2 = cad2 & Cad & DBSet(NumF, "N") & "," & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "N") & ","
        cad2 = cad2 & DBLet(Rs.Fields(6).Value, "N") & "," 'cantreal <- pesoneto
        cad2 = cad2 & DBSet(DBLet(Rs.Fields(4).Value, "N") * DBLet(Rs.Fields(10).Value, "N"), "N") & "," 'cantfact <- numcajas * kiloscaj
        'cad2 = cad2 & DBSet(RS.Fields(7).Value, "N") & "," 'preciobru <- preciopro
        '[Monica]23/04/2013: cogemos el precio definitivo o el provisional
        cad2 = cad2 & DBSet(vPrecio, "N") & "," 'preciobru <- preciopro
        cad2 = cad2 & "0," ' precinet
        cad2 = cad2 & DBSet(vCliente.Dto1, "N") & "," ' dtocom1
        cad2 = cad2 & DBSet(vCliente.Dto2, "N") & "," ' dtocom2
        
        '++monica:27/05/08 dependiendo de si la confeccion es por unidades(0) o por kilos(1)
        If DBLet(Rs.Fields(13).Value, "N") = 0 Then 'unidades
            'antes:                                                      DBLet(RS.Fields(7).Value, "N"), 2), "N") & "," ' imporbru
            cad2 = cad2 & DBSet(Round2(DBLet(Rs.Fields(11).Value, "N") * DBLet(vPrecio, "N"), 2), "N") & "," ' imporbru
        Else   'kilos
            'antes:                                                      DBLet(RS.Fields(7).Value, "N"), 2), "N") & "," ' imporbru
            cad2 = cad2 & DBSet(Round2(DBLet(Rs.Fields(6).Value, "N") * DBLet(vPrecio, "N"), 2), "N") & "," ' imporbru
        End If
        
        cad2 = cad2 & "0," ' impornet
        cad2 = cad2 & DBLet(TipoIva, "N") & "," ' codigiva
        cad2 = cad2 & DBSet(Rs.Fields(11).Value, "N") & "),"
        
        NumF = NumF + 1
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    ' quitamos la ultima ,
    cad2 = Mid(cad2, 1, Len(cad2) - 1)
    cad2 = Sql & cad2
    
    conn.Execute cad2

    InsertarLineasFactu = True
    
EInsertarLin:
    If Err.Number <> 0 Then
        InsertarLineasFactu = False
        MuestraError Err.Number, "Insertar lineas factura.", Err.Description
    Else
        InsertarLineasFactu = True
    End If
End Function


Private Function ClienteExentoIVA() As Boolean
Dim vClien As CCliente
Dim b As Boolean

    Set vClien = New CCliente
    vClien.Codigo = Cliente
    b = vClien.exentoIVA
    Set vClien = Nothing
    
    ClienteExentoIVA = b
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
    If vParamAplic.ContabilidadNueva Then
        cadAux = DevuelveDesdeBDNew(cConta, "formapago", "codforpa", "codforpa", nForPa, "N")
    Else
        cadAux = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", nForPa, "N")
    End If
    'si no existe la forma de pago en conta, insertamos la de ariges
    If cadAux = "" Then
        cadAux2 = "tipforpa"
        cadAux = DevuelveDesdeBDNew(cAgro, "forpago", "nomforpa", "codforpa", nForPa, "N", cadAux2)
        If cadAux <> "" Then
            'insertamos e sforpa de la CONTA
            If vParamAplic.ContabilidadNueva Then
                Sql = "INSERT INTO formapago(codforpa,nomforpa,tipforpa)"
                Sql = Sql & " VALUES(" & nForPa & ", " & DBSet(cadAux, "T") & ", " & cadAux2 & ")"
            Else
                Sql = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
                Sql = Sql & " VALUES(" & nForPa & ", " & DBSet(cadAux, "T") & ", " & cadAux2 & ")"
            End If
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



Private Function RecalcularDtos(ByRef vCliente As CCliente) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String
Dim vImpDto As Currency
Dim vDto1 As Currency
Dim vDto2 As Currency
Dim vImpNeto As Currency
Dim vPrecNeto As Currency
Dim TipoDto As String
Dim ImpDto As String
Dim Rdc As Long

    On Error GoTo eRecalcularDtos

    Sql = "select * from facturas_variedad where codtipom = " & DBSet(codTipoM, "T")
    Sql = Sql & " and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
    Sql = Sql & " order by numlinea "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    vImpDto = 0
    
    vDto1 = ComprobarCero(vCliente.Dto1)
    
    vDto2 = ComprobarCero(vCliente.Dto2)
    '++monica:030608: traemos el redondeo de clientes
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "clientes", "nrodecprec", "codclien", vCliente.Codigo, "N")
    Rdc = ComprobarCero(Sql)
    '++
    While Not Rs.EOF
        Select Case TipoFacturarForfaits(CStr(Rs!NumAlbar), CStr(Rs!numlinealbar))
            Case 0  'unidades
                ImpDto = CalcularImporteDto(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), CStr(vImpDto), True)
                vImpNeto = CalcularImporte(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), CStr(vDto1), CStr(vDto2), CByte(vCliente.TipoDto), CStr(ImpDto))
                
                'precio neto
                '[Monica]24/11/2011: añadida la condicion para que no dé division por cero
                vPrecNeto = 0
                If DBLet(Rs!Unidades, "N") <> 0 Then
                    vPrecNeto = Round2(vImpNeto / DBLet(Rs!Unidades, "N"), Rdc)
                End If
                If Rdc = 2 Or Rdc = 3 Then
                    vImpNeto = Round2(vPrecNeto * DBLet(Rs!Unidades, "N"), 2)
                End If
            Case 1  'kilos
                ImpDto = CalcularImporteDto(DBLet(Rs!cantreal, "N"), DBLet(Rs!precibru, "N"), CStr(vImpDto), True)
                vImpNeto = CalcularImporte(DBLet(Rs!cantreal, "N"), DBLet(Rs!precibru, "N"), CStr(vDto1), CStr(vDto2), CByte(vCliente.TipoDto), CStr(ImpDto))
                
                'precio neto
                '[Monica]24/11/2011: añadida la condicion para que no dé division por cero
                vPrecNeto = 0
                If DBLet(Rs!cantreal, "N") <> 0 Then
                    vPrecNeto = Round2(vImpNeto / DBLet(Rs!cantreal, "N"), Rdc)
                End If
                If Rdc = 2 Or Rdc = 3 Then
                    vImpNeto = Round2(vPrecNeto * DBLet(Rs!cantreal, "N"), 2)
                End If
            Case Else
        
        End Select
        Sql2 = "update facturas_variedad set impornet = " & DBSet(vImpNeto, "N")
        Sql2 = Sql2 & ",precinet = " & DBSet(vPrecNeto, "N")
        Sql2 = Sql2 & " where codtipom = " & DBSet(codTipoM, "T")
        Sql2 = Sql2 & " and numfactu = " & DBSet(NumFactu, "N")
        Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
        Sql2 = Sql2 & " and numlinea = " & DBSet(Rs!NumLinea, "N")
    
        conn.Execute Sql2
    
        Rs.MoveNext
    Wend

    Set Rs = Nothing
    
    RecalcularDtos = True
    Exit Function

eRecalcularDtos:
    If Err.Number <> 0 Then
        RecalcularDtos = False
    End If

End Function

Private Function CalcularImporteDto(Cantidad As String, Precio As String, ImpDto As String, Insertado As Boolean) As String
'Insertado: indica si ya hemos insertado el registro o no
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
Dim vCant As Currency
Dim vImp As Currency
Dim vDto As Currency
Dim vPre As Currency
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim SumaBruto As Currency

On Error Resume Next


    Sql = "select sum(imporbru) from facturas_variedad where codtipom = " & DBSet(codTipoM, "T")
    Sql = Sql & " and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    SumaBruto = 0
    If Not Rs.EOF Then
        SumaBruto = DBLet(Rs.Fields(0).Value, "N")
    End If
    
    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    vCant = ComprobarCero(Cantidad)
    vPre = ComprobarCero(Precio)
    vDto = ComprobarCero(ImpDto)
    
    If Not Insertado Then
        SumaBruto = SumaBruto + (CCur(vCant) * CCur(vPre))
    End If
    
    If SumaBruto <> 0 Then
        vImp = (CCur(Cantidad) * CCur(vPre) * CCur(vDto)) / SumaBruto
    Else
        vImp = CCur(vDto)
    End If
    
    vImp = Round2(vImp, 6)
    
    CalcularImporteDto = CStr(vImp)

End Function




Public Function PasarAlbaranesAFactura2(TipoAlb As String, cadSQL As String, textoCSB As String, Optional ErroresAux As String, Optional EstaRecu As Boolean) As Boolean
'IN ->  cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'       TextosCSB:  Para la tesoreria
'Desde Albaranes Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETraspasoAlbFac

    PasarAlbaranesAFactura2 = False
    b = False
    ErroresAux = ""
    '#Laura 14/11/2006 recuperar facturas de Alzira
    'si esta recuperando facturas
    EstaRecuFac = EstaRecu
    
    Select Case TipoAlb
        Case "ALV", "AL1" 'ALV: Albaranes venta envases a clientes
            codTipoM = "FAV"
    End Select
      
     
    'Aqui empieza transaccion
    conn.BeginTrans
   
    'Insertar la Factura (scafac,scafac1,slifac)
    MenError = "Error al Insertar en tablas de factura."
    b = InsertarFacturaEnvases(cadSQL)
    
    
    'Eliminar los Albaranes incluidos en la Factura
    If b Then
        MenError = "Error al eliminar los albaranes."
        b = EliminarAlbaranesEnvases(cadSQL)
    End If
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        ErroresAux = Err.Description
        Err.Clear
        b = False
    End If
  
    If b Then
        conn.CommitTrans
        PasarAlbaranesAFactura2 = True
    Else
        conn.RollbackTrans
        PasarAlbaranesAFactura2 = False
        ErroresAux = MenError & vbCrLf & ErroresAux & vbCrLf
    End If
    espera 0.4
End Function




Private Function InsertarFacturaEnvases(cadSQL As String) As Boolean
'Insertamos en las tablas de factura de cliente: facturas, facturas_envases
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean
Dim devuelve As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim vCliente As CCliente
Dim TipoMovimiento As String

    InsertarFacturaEnvases = False
    
    'cogemos los datos del albaran
    Sql = "select codclien, fechaalb, observac from scaalb where " & cadSQL
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Set vCliente = New CCliente
    
    'si se ha modificado el cliente volver a cargar los datos
    If vCliente.Existe(Rs.Fields(0).Value) Then
        If vCliente.LeerDatos(Rs.Fields(0).Value) Then
             Cliente = vCliente.Codigo
            
            If vCliente.TipoFactu = 0 Then
                ' 0 = automatica 1 = manual
                TipoMovimiento = vCliente.tipoMov
                Set vTipoMov = New CTiposMov
                If vTipoMov.Leer(TipoMovimiento) Then
                    NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                    ' si existe la factura incrementamos el contador
                    Do
                        devuelve = DevuelveDesdeBDNew(cAgro, "facturas", "numfactu", "codtipom", TipoMovimiento, "T", , "numfactu", NumFactu, "N", "fecfactu", Rs.Fields(1).Value, "F")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (TipoMovimiento)
                            NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                Else
                    Exit Function
                End If
            Else
                ' pedimos el tipo de movimiento y el numero de factura
'--monica:040608
'                TipoMovimiento = "FAV"
'                NumFactu = InputBox("Número de factura:")
'                If MsgBox("¿Datos correctos para continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                    Exit Function
'                End If
'                devuelve = DevuelveDesdeBDNew(cAgro, "facturas", "numfactu", "codtipom", TipoMovimiento, "T", , "numfactu", NumFactu, "N", "fecfactu", Rs.Fields(1).Value, "F")
'                If devuelve <> "" Then
'                    MsgBox "Factura existente. Revise.", vbExclamation
'                    Exit Function
'                End If
'++monica:040608: pedimos unicamente el tipo de movimiento
                TipoMovimiento = "FAV" 'InputBox("Tipo Factura:")
'                If MsgBox("¿Datos correctos para continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                    Exit Function
'                End If
                Set vTipoMov = New CTiposMov
                If vTipoMov.Leer(TipoMovimiento) Then
                    NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                    ' si existe la factura incrementamos el contador
                    Do
                        devuelve = DevuelveDesdeBDNew(cAgro, "facturas", "numfactu", "codtipom", TipoMovimiento, "T", , "numfactu", NumFactu, "N", "fecfactu", mFecFactu, "F")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (TipoMovimiento)
                            NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                Else
                    Exit Function
                End If
            
            End If
            
            codTipoM = TipoMovimiento
            
            'Cabecera Factura
            b = InsertarCabeceraFactu(vCliente, DBLet(Rs.Fields(2).Value, "T"))
            If Not b Then Exit Function
             
            'Insertar lineas de variedades de la factura
            b = InsertarLineasFactuEnv(cadSQL)
            If Not b Then Exit Function
                         
            'calcular el total de la factura
            b = CalcularDatosFacturaVenta(vCliente)
            If Not b Then Exit Function
            
            'Incrementar contador del tipo de movimiento
            '-------------------------------------------------------------
'--monica: en cualquier caso se incrementa contador
'            If vCliente.TipoFactu = 0 Then
                vTipoMov.IncrementarContador (codTipoM)
                Set vTipoMov = Nothing
'            End If
            
            InsertarFacturaEnvases = True
            
        End If
    End If
    Set vCliente = Nothing
    
End Function


Private Function EliminarAlbaranesEnvases(cadSQL As String) As Boolean
'Eliminamos de las tablas de Albaranes: scaalb, slialb
Dim Sql As String

    On Error GoTo EEliminar

    EliminarAlbaranesEnvases = False
    
    'ELiminar lineas albaranes
    Sql = "DELETE FROM slialb "
    Sql = Sql & " WHERE " & Replace(cadSQL, "scaalb", "slialb")
    conn.Execute Sql
    
    espera 0.1
    
    'Eliminar Cabeceras Albaranes
    Sql = "DELETE FROM scaalb "
    Sql = Sql & " WHERE " & Replace(cadSQL, "slialb", "scaalb")
    conn.Execute Sql
        
    EliminarAlbaranesEnvases = True

EEliminar:
    If Err.Number <> 0 Then
        EliminarAlbaranesEnvases = False
    Else
        EliminarAlbaranesEnvases = True
    End If
End Function


Private Function InsertarLineasFactuEnv(cadSQL As String) As Boolean
'Insertar las lineas de los Albaranes de envases de la factura en la tabla: slifac
Dim Sql As String
Dim Sql2 As String
Dim Cad As String
Dim cad2 As String
Dim NumF As Long
Dim Rs As ADODB.Recordset
Dim FecAlbar As String

    On Error GoTo EInsertarLin

    Sql = "INSERT INTO facturas_envases (codtipom,numfactu,fecfactu,numlinea, codalmac,codartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva,numalbar,numlinealbar,fecalbar) values "
    
     'pasamos lineas de albaranes a lineas factura
    Sql2 = " SELECT numalbar,numlinea,codalmac,slialb.codartic,cantidad,precioar,dtolinea,importel, ampliaci, codigiva  "
    Sql2 = Sql2 & " FROM slialb "
    Sql2 = Sql2 & " WHERE " & Replace(cadSQL, "scaalb", "slialb")
    

    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Cad = "(" & DBSet(codTipoM, "T") & "," & DBSet(NumFactu, "N") & "," & DBSet(FecFactu, "F") & ","
    cad2 = ""
    
    NumF = 1
    
    FecAlbar = DevuelveValor("select fechaalb from scaalb where " & cadSQL)
    
    While Not Rs.EOF
        cad2 = cad2 & Cad & DBSet(NumF, "N") & "," & DBSet(Rs.Fields(2).Value, "N") & "," & DBSet(Rs.Fields(3).Value, "T") & ","
        cad2 = cad2 & DBSet(Rs.Fields(4).Value, "N") & "," 'cantidad
        cad2 = cad2 & DBSet(Rs.Fields(5).Value, "N") & "," 'precio
        cad2 = cad2 & DBSet(Rs.Fields(6).Value, "N") & "," 'dto
        cad2 = cad2 & DBSet(Rs.Fields(7).Value, "N") & "," 'importe
        cad2 = cad2 & DBSet(Rs.Fields(8).Value, "T", "S") & "," 'ampliacion
        cad2 = cad2 & DBSet(Rs.Fields(9).Value, "N") & "," 'codigiva
        cad2 = cad2 & DBSet(Rs.Fields(0).Value, "N") & "," 'numalbar
        cad2 = cad2 & DBSet(Rs.Fields(1).Value, "N") & "," 'numlinea
        cad2 = cad2 & DBSet(FecAlbar, "F") & ")," ' fecha de albaran que es la fecha de movimiento
        
        NumF = NumF + 1
        
        Rs.MoveNext
    Wend
    
    ' quitamos la ultima coma
    If cad2 <> "" Then
        cad2 = Mid(cad2, 1, Len(cad2) - 1)
    
        conn.Execute Sql & cad2
    End If

    InsertarLineasFactuEnv = True
    
EInsertarLin:
    If Err.Number <> 0 Then
        InsertarLineasFactuEnv = False
        MuestraError Err.Number, "Insertar lineas factura envases.", Err.Description
    Else
        InsertarLineasFactuEnv = True
    End If
End Function

'#############################################################################################################
'################################  FACTURAS DE VENTA A SOCIOS   ##############################################
'#############################################################################################################


Public Function PasarAlbaranAFacturaSOCIO(cadSQL As String, FechaFact As String, Optional ErroresAux As String, Optional EstaRecu As Boolean) As Boolean
'IN -> cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'Desde Albaranes Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETraspasoAlbFac

    PasarAlbaranAFacturaSOCIO = False
    b = False
    ErroresAux = ""
      
    'Aqui empieza transaccion
    conn.BeginTrans
    ConnConta.BeginTrans
   
    FecFactu = FechaFact
   
    'Insertar la Factura (scafac,scafac1,slifac)
    MenError = "Error al Insertar en tablas de factura."
    b = InsertarFacturaSOCIO(cadSQL)
    
    
'    'Insertar los Vencimientos de la Factura. Tabla: svenci
'    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
'    '#Laura: 14/11/2006 Recuperar facturas Alzira
'    If b Then
'        If EstaRecuFac = False Then
'            MenError = "Error al actualizar en Tesoreria."
'            b = InsertarEnTesoreria(MenError)
'        Else
'            MenError = "Error al actualizar en Tesoreria."
'            b = InsertarEnTesoreria_RecupFac(MenError)
'        End If
'    End If
'
    
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        ErroresAux = Err.Description
        Err.Clear
        b = False
    End If
  
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        PasarAlbaranAFacturaSOCIO = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        PasarAlbaranAFacturaSOCIO = False
        ErroresAux = MenError & vbCrLf & ErroresAux & vbCrLf
    End If
    espera 0.4
End Function


Private Function InsertarFacturaSOCIO(cadSQL As String) As Boolean
'Insertamos en las tablas de factura de socio: scafac, scafac1, slifac
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean
Dim devuelve As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim vSocio As CSocio
Dim TipoMovimiento As String
Dim NumFactuNue As String
Dim NumFactuAnt As String
Dim SeccionHorto As String

    InsertarFacturaSOCIO = False
    
    SeccionHorto = DevuelveValor("select seccionhorto from rparam")
    
    'cogemos los datos del albaran
    Sql = "select codsocio, fechaalb, observac from albaran where " & cadSQL
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Set vSocio = New CSocio
    
    'si se ha modificado el cliente volver a cargar los datos
    If vSocio.Existe(Rs.Fields(0).Value) Then
        If vSocio.LeerDatos(Rs.Fields(0).Value) Then
            If vSocio.LeerDatosSeccion(vSocio.Codigo, SeccionHorto) Then
                Socio = vSocio.Codigo
                ' 0 = automatica 1 = manual
                TipoMovimiento = "FAS"
                Set vTipoMov = New CTiposMov
                If vTipoMov.Leer(TipoMovimiento) Then
                    NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                    NumFactuAnt = NumFactu
                    ' si existe la factura incrementamos el contador
                    Do
                        devuelve = DevuelveDesdeBDNew(cAgro, "facturassocio", "numfactu", "codtipom", TipoMovimiento, "T", , "numfactu", NumFactu, "N", "fecfactu", FecFactu, "F")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (TipoMovimiento)
                            NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                Else
                    Exit Function
                End If
                
                codTipoM = TipoMovimiento
                
                'Cabecera Factura
                b = InsertarCabeceraFactuSOCIO(vSocio, DBLet(Rs.Fields(2).Value, "T"))
                If Not b Then Exit Function
                 
                'Insertar lineas de variedades de la factura
                b = InsertarLineasFactuSOCIO(cadSQL, vSocio)
                If Not b Then Exit Function
                             
                b = RecalcularDtosSOCIO(vSocio)
                If Not b Then Exit Function
                
                'calcular el total de la factura
                b = CalcularDatosFacturaVentaSOCIO(vSocio)
                If Not b Then Exit Function
                
                'Incrementar contador del tipo de movimiento
                '-------------------------------------------------------------
    '--monica: en cualquier caso se incrementa contador
    '            If vCliente.TipoFactu = 0 Then
                If NumFactuAnt = NumFactu Then
                    vTipoMov.IncrementarContador (codTipoM)
                    Set vTipoMov = Nothing
                End If
            
            InsertarFacturaSOCIO = True
            End If
        End If
    End If
    Set vSocio = Nothing
    
End Function

Private Function InsertarCabeceraFactuSOCIO(ByRef vSocio As CSocio, Observac As String) As Boolean
'Inserta la cabecera de la factura en la tabla: scafac
Dim Sql As String
Dim vClien As CCliente
Dim Nulo2 As String
Dim Nulo3 As String
Dim ctaDpto As Boolean
Dim Cad As String, cad2 As String

    On Error GoTo EInsertar


    'Insertar en la tabla cabecera de la factura de ventas
    Sql = "INSERT INTO facturassocio (codtipom,numfactu,fecfactu,codsocio,codforpa,tipoivac,dtocom1,"
    Sql = Sql & "dtocom2,impdtoc,observac,baseimp1,baseimp2,baseimp3,impoiva1,impoiva2,impoiva3,"
    Sql = Sql & "imporec1,imporec2,imporec3,codiiva1,codiiva2,codiiva3,porciva1,porciva2,porciva3,"
    Sql = Sql & "porcrec1,porcrec2,porcrec3,brutofac,impordto,totalfac,intconta,pasaridoc) "
    
    Sql = Sql & " VALUES (" & DBSet(codTipoM, "T") & "," & DBSet(NumFactu, "N") & "," & DBSet(FecFactu, "F") & "," & DBSet(vSocio.Codigo, "N") & ","
    Sql = Sql & DBSet(vSocio.ForPago, "N") & "," & DBSet(vSocio.Codiva, "N") & "," & "0,0," 'DBSet(vSocio.Dto1, "N") & "," & DBSet(vSocio.Dto2, "N") & ","
    Sql = Sql & ValorNulo & "," & DBSet(Observac, "T", "S") & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & "0,0)"
    
    conn.Execute Sql
    
    InsertarCabeceraFactuSOCIO = True
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabeceraFactuSOCIO = False
        MuestraError Err.Number, "Insertar cabecera factura SOCIO.", Err.Description
    Else
        InsertarCabeceraFactuSOCIO = True
    End If
End Function


Private Function InsertarLineasFactuSOCIO(vSQL As String, ByRef vSocio As CSocio) As Boolean
'Insertar las lineas de los Albaranes de la factura en la tabla: facturassocio_variedad
Dim Sql As String
Dim Sql2 As String
Dim Cad As String
Dim cad2 As String
Dim vWhere2 As String
Dim cArt As CArticulo
Dim Rs As ADODB.Recordset
Dim NumF As Long
Dim TipoIva As Integer
Dim vSql1 As String

    On Error GoTo EInsertarLin

    Sql = "INSERT INTO facturassocio_variedad (codtipom,numfactu,fecfactu,numlinea,numalbar,"
    Sql = Sql & "numlinealbar,cantreal,cantfact,precibru,precinet,dtocom1,dtocom2,imporbru,"
    Sql = Sql & "impornet,codigiva,unidades) values "
    
    vWhere2 = "codtipom = " & DBSet(codTipoM, "T") & " and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
    
'    Select Case vSocio.TipoIva
'        Case 0  'normal
'            TipoIva = vParamAplic.CodIvaNormal
'++monica:27/05/08 el tipo de iva lo sacaré de la variedad
            TipoIva = 0
'        Case 1  'exento
'            TipoIva = vParamAplic.CodIvaExento
'        Case 2  'recargo de equivalencia
'            TipoIva = vParamAplic.CodIvaRecargo
'    End Select
    
    vSql1 = Replace(vSQL, "albaran", "albaran_variedad")
    
    Sql2 = "select albaran_variedad.numalbar, albaran_variedad.numlinea, albaran_variedad.codforfait,"
    Sql2 = Sql2 & " albaran_variedad.totpalet, albaran_variedad.numcajas, albaran_variedad.pesobrut,"
    Sql2 = Sql2 & " albaran_variedad.pesoneto, albaran_variedad.preciopro, albaran_variedad.preciodef,"
    Sql2 = Sql2 & " albaran_variedad.impcomis, forfaits.kiloscaj, albaran_variedad.unidades, "
    Sql2 = Sql2 & " albaran_variedad.codvarie, forfaits.facturar "
    Sql2 = Sql2 & " from albaran_variedad, forfaits where " & vSql1
    Sql2 = Sql2 & " and albaran_variedad.codforfait = forfaits.codforfait "
    Sql2 = Sql2 & " order by albaran_variedad.numalbar, albaran_variedad.numlinea"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Cad = "(" & DBSet(codTipoM, "T") & "," & DBSet(NumFactu, "N") & "," & DBSet(FecFactu, "F") & ","
    cad2 = ""
    
    NumF = 1
    
    While Not Rs.EOF
'        If vSocio.TipoIva = 0 Then
            TipoIva = DevuelveDesdeBDNew(cAgro, "variedades", "codigiva", "codvarie", Rs.Fields(12).Value, "N")
'        End If
        cad2 = cad2 & Cad & DBSet(NumF, "N") & "," & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "N") & ","
        cad2 = cad2 & DBSet(Rs.Fields(6).Value, "N") & "," 'cantreal <- pesoneto
        cad2 = cad2 & DBSet(DBLet(Rs.Fields(4).Value, "N") * DBLet(Rs.Fields(10).Value, "N"), "N") & "," 'cantfact <- numcajas * kiloscaj
        cad2 = cad2 & DBSet(DBLet(Rs.Fields(7).Value, "N"), "N") & "," 'preciobru <- preciopro
        cad2 = cad2 & "0," ' precinet
        cad2 = cad2 & "0," 'DBSet(vSocio.Dto1, "N") & "," ' dtocom1
        cad2 = cad2 & "0," 'DBSet(vSocio.Dto2, "N") & "," ' dtocom2
        
        '++monica:27/05/08 dependiendo de si la confeccion es por unidades(0) o por kilos(1)
        If DBLet(Rs.Fields(13).Value, "N") = 0 Then 'unidades
            cad2 = cad2 & DBSet(Round2(Rs.Fields(11).Value * DBLet(Rs.Fields(7).Value, "N"), 2), "N") & "," ' imporbru
        Else   'kilos
            cad2 = cad2 & DBSet(Round2(Rs.Fields(6).Value * DBLet(Rs.Fields(7).Value, "N"), 2), "N") & "," ' imporbru
        End If
        
        cad2 = cad2 & "0," ' impornet
        cad2 = cad2 & DBSet(TipoIva, "N") & "," ' codigiva
        cad2 = cad2 & DBSet(Rs.Fields(11).Value, "N") & "),"
        
        NumF = NumF + 1
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    ' quitamos la ultima ,
    cad2 = Mid(cad2, 1, Len(cad2) - 1)
    cad2 = Sql & cad2
    
    conn.Execute cad2

    InsertarLineasFactuSOCIO = True
    
EInsertarLin:
    If Err.Number <> 0 Then
        InsertarLineasFactuSOCIO = False
        MuestraError Err.Number, "Insertar lineas factura Socio.", Err.Description
    Else
        InsertarLineasFactuSOCIO = True
    End If
End Function


Private Function CalcularDatosFacturaVentaSOCIO(vSocio As CSocio) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim Rs As ADODB.Recordset
Dim i As Integer

Dim Sql As String
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

Dim cadwhere As String
Dim cadwhere1 As String
    
Dim Nulo2 As String
Dim Nulo3 As String

    
    CalcularDatosFacturaVentaSOCIO = False
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
    cadwhere = "facturassocio.codtipom = " & DBSet(codTipoM, "T") & " and facturassocio.numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
    cadwhere1 = Replace(cadwhere, "facturassocio", "facturassocio_variedad")
    
    Sql = "SELECT facturassocio_variedad.codigiva, sum(imporbru) as bruto, sum(impornet) as neto"
    Sql = Sql & " FROM facturassocio_variedad "
    Sql = Sql & " WHERE " & cadwhere1
    Sql = Sql & " GROUP BY 1 "
    Sql = Sql & " UNION "
    cadwhere1 = Replace(cadwhere, "facturassocio", "facturassocio_envases")
    Sql = Sql & "SELECT facturassocio_envases.codigiva, sum(importel) as bruto, sum(importel) as neto"
    Sql = Sql & " FROM facturassocio_envases "
    Sql = Sql & " WHERE " & cadwhere1
    Sql = Sql & " GROUP BY 1 "
    Sql = Sql & " ORDER BY 1 "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

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

'            If vSocio.TipoIva = 2 Then
'                'Obtener el % de RECARGO
'                cadAux1 = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", CStr(IvaAnt), "N")
'
'                'aplicar el RECARGO a la base imponible de ese tipo
'                ImpREC = CalcularPorcentaje(ImpBImIVA, CCur(cadAux1), 2)
'
'                'sumamos todos los RECARGOS para sumarselo a la base imponible total de la factura
'                'los vamos acumulando
'                TotImpIVA = TotImpIVA + ImpREC
'            Else
                cadAux1 = "0"
                ImpREC = 0
'            End If


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
    
'    If CInt(vSocio.TipoIva) = 2 Then
'        'Obtener el % de RECARGO
'        cadAux1 = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", CStr(IvaAnt), "N")
'
'        'aplicar el RECARGO a la base imponible de ese tipo
'        ImpREC = CalcularPorcentaje(ImpBImIVA, CCur(cadAux1), 2)
'    Else
        cadAux1 = "0"
        ImpREC = 0
'    End If
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

    Nulo2 = "N"
    Nulo3 = "N"
    If DBSet(TipoIVA2, "N", "S") = ValorNulo Then Nulo2 = "S"
    If DBSet(TipoIVA3, "N", "S") = ValorNulo Then Nulo3 = "S"
    
    'ACTUALIZAMOS LA FACTURA (tabla facturas)
    Sql = "update facturassocio "
    Sql = Sql & "set baseimp1 = " & DBSet(BaseIVA1, "N")
    Sql = Sql & ",baseimp2 = " & DBSet(BaseIVA2, "N", Nulo2)
    Sql = Sql & ",baseimp3 = " & DBSet(BaseIVA3, "N", Nulo3)
    Sql = Sql & ",impoiva1 = " & DBSet(ImpIVA1, "N")
    Sql = Sql & ",impoiva2 = " & DBSet(ImpIVA2, "N", Nulo2)
    Sql = Sql & ",impoiva3 = " & DBSet(ImpIVA3, "N", Nulo3)
    Sql = Sql & ",imporec1 = " & DBSet(ImpREC1, "N")
    Sql = Sql & ",imporec2 = " & DBSet(ImpREC2, "N", Nulo2)
    Sql = Sql & ",imporec3 = " & DBSet(ImpREC3, "N", Nulo3)
    Sql = Sql & ",codiiva1 = " & DBSet(TipoIVA1, "N")
    Sql = Sql & ",codiiva2 = " & DBSet(TipoIVA2, "N", Nulo2)
    Sql = Sql & ",codiiva3 = " & DBSet(TipoIVA3, "N", Nulo3)
    Sql = Sql & ",porciva1 = " & DBSet(PorceIVA1, "N")
    Sql = Sql & ",porciva2 = " & DBSet(PorceIVA2, "N", Nulo2)
    Sql = Sql & ",porciva3 = " & DBSet(PorceIVA3, "N", Nulo3)
    Sql = Sql & ",porcrec1 = " & DBSet(PorceREC1, "N")
    Sql = Sql & ",porcrec2 = " & DBSet(PorceREC2, "N", Nulo2)
    Sql = Sql & ",porcrec3 = " & DBSet(PorceREC3, "N", Nulo3)
    Sql = Sql & ",brutofac = " & DBSet(TotBruto, "N")
    Sql = Sql & ",impordto = " & DBSet(Round2(TotBruto - TotNeto, 2), "N")
    Sql = Sql & ",totalfac = " & DBSet(TotalFac, "N")
    Sql = Sql & " where " & cadwhere
    
    conn.Execute Sql

    CalcularDatosFacturaVentaSOCIO = True

ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFacturaVentaSOCIO = False
    Else
        CalcularDatosFacturaVentaSOCIO = True
    End If
End Function


Private Function RecalcularDtosSOCIO(ByRef vSocio As CSocio) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String
Dim vImpDto As Currency
Dim vDto1 As Currency
Dim vDto2 As Currency
Dim vImpNeto As Currency
Dim vPrecNeto As Currency
Dim TipoDto As String
Dim ImpDto As String
Dim Rdc As Long

    On Error GoTo eRecalcularDtosSOCIO

    Sql = "select * from facturassocio_variedad where codtipom = " & DBSet(codTipoM, "T")
    Sql = Sql & " and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
    Sql = Sql & " order by numlinea "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    vImpDto = 0
    
    vDto1 = 0 'ComprobarCero(vSocio.Dto1)
    
    vDto2 = 0 'ComprobarCero(vSocio.Dto2)
    '++monica:030608: traemos el redondeo de clientes
'    Sql = ""
'    Sql = DevuelveDesdeBDNew(cAgro, "clientes", "nrodecprec", "codclien", vCliente.Codigo, "N")
    Rdc = 4 'ComprobarCero(Sql)
    '++
    While Not Rs.EOF
        Select Case TipoFacturarForfaits(CStr(Rs!NumAlbar), CStr(Rs!numlinealbar))
            Case 0  'unidades
                ImpDto = CalcularImporteDto(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), CStr(vImpDto), True)
                vImpNeto = CalcularImporte(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), CStr(vDto1), CStr(vDto2), 0, CStr(ImpDto)) 'CByte(vSocio.TipoDto), CStr(ImpDto))
                
                'precio neto
                vPrecNeto = Round2(vImpNeto / DBLet(Rs!Unidades, "N"), Rdc)
                If Rdc = 2 Or Rdc = 3 Then
                    vImpNeto = Round2(vPrecNeto * DBLet(Rs!Unidades, "N"), 2)
                End If
            Case 1  'kilos
                ImpDto = CalcularImporteDto(DBLet(Rs!cantreal, "N"), DBLet(Rs!precibru, "N"), CStr(vImpDto), True)
                vImpNeto = CalcularImporte(DBLet(Rs!cantreal, "N"), DBLet(Rs!precibru, "N"), CStr(vDto1), CStr(vDto2), 0, CStr(ImpDto)) 'CByte(vCliente.TipoDto), CStr(ImpDto))
                
                'precio neto
                vPrecNeto = Round2(vImpNeto / DBLet(Rs!cantreal, "N"), Rdc)
                If Rdc = 2 Or Rdc = 3 Then
                    vImpNeto = Round2(vPrecNeto * DBLet(Rs!cantreal, "N"), 2)
                End If
            Case Else
        
        End Select
        Sql2 = "update facturassocio_variedad set impornet = " & DBSet(vImpNeto, "N")
        Sql2 = Sql2 & ",precinet = " & DBSet(vPrecNeto, "N")
        Sql2 = Sql2 & " where codtipom = " & DBSet(codTipoM, "T")
        Sql2 = Sql2 & " and numfactu = " & DBSet(NumFactu, "N")
        Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
        Sql2 = Sql2 & " and numlinea = " & DBSet(Rs!NumLinea, "N")
    
        conn.Execute Sql2
    
        Rs.MoveNext
    Wend

    Set Rs = Nothing
    
    RecalcularDtosSOCIO = True
    Exit Function

eRecalcularDtosSOCIO:
    If Err.Number <> 0 Then
        RecalcularDtosSOCIO = False
    End If

End Function


'[Monica]25/10/2011: Añadido
Private Function DescontarFacturasACuenta(TipoM As String, Factu As String, fecFac As String, Cliente As String) As Boolean
Dim Sql As String
Dim cadwhere As String


    On Error GoTo eDescontarFacturasACuenta

    DescontarFacturasACuenta = False


    Sql = "select codtipom, numfactu, fecfactu, totalfac from facturas "
    cadwhere = "where codtipom = 'EAC' and codclien = " & DBSet(Cliente, "N")
    cadwhere = cadwhere & " and (codtipom, numfactu, fecfactu) not in (select codtipomcta, numfactucta, fecfactucta from facturas_acuenta) "
'    cadWHERE = cadWHERE & " where codclien = " & DBSet(Cliente, "N") & " and codtipom = 'EAC')"
    
    Sql = Sql & cadwhere
    
    If TotalRegistrosConsulta(Sql) <> 0 Then
        
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 22
        frmMens.cadwhere = cadwhere
        
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
'    Else
'        MsgBox "No hay facturas a cuenta para descontar.", vbExclamation
    End If

    DescontarFacturasACuenta = True
    Exit Function
    
eDescontarFacturasACuenta:
    MuestraError Err.Number, "Descontar Facturas a Cuenta", Err.Description
End Function

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        FacturasADescontar = " (codtipom, numfactu, fecfactu) in (" & CadenaSeleccion & ")"
    Else
        FacturasADescontar = ""
    End If
  
    If FacturasADescontar <> "" Then
        Sql = "insert into facturas_acuenta (codtipom,numfactu,fecfactu,codtipomcta,numfactucta,fecfactucta,totalfaccta)"
        Sql = Sql & " select " & DBSet(codTipoM, "T") & "," & DBSet(NumFactu, "N") & "," & DBSet(FecFactu, "F") & ", codtipom, numfactu, fecfactu, totalfac from facturas "
        Sql = Sql & " where " & FacturasADescontar
        
        conn.Execute Sql
        
    End If
  
End Sub


Private Function InsertarLineasCalibreFactu(vSQL As String) As Boolean
'Insertar las lineas de calibre de los Albaranes de la factura en la tabla: facturas_variedad
Dim Sql2 As String
Dim vWhere2 As String
Dim Rs As ADODB.Recordset
Dim NumF As Long
Dim b As Boolean
Dim Mens As String

    On Error GoTo EInsertarLin

    
    vWhere2 = "codtipom = " & DBSet(codTipoM, "T") & " and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
    
    
    
    Sql2 = "select codtipom,numfactu,fecfactu,numlinea,numalbar,numlinealbar,cantreal,cantfact,precibru,precinet,dtocom1,dtocom2,imporbru,impornet,codigiva,unidades "
    Sql2 = Sql2 & " from facturas_variedad where " & vWhere2
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    NumF = 1
    b = True
    
    While Not Rs.EOF And b
    
        Mens = "Insertar Calibres"
        If b Then b = InsertarModificarCalibres(True, codTipoM, NumFactu, FecFactu, CStr(DBLet(Rs!NumLinea, "N")), CStr(DBLet(Rs!NumAlbar, "N")), CStr(DBLet(Rs!numlinealbar, "N")), CStr(DBLet(Rs!cantreal, "N")), CStr(DBLet(Rs!Unidades, "N")), CStr(DBLet(Rs!imporbru, "N")), CStr(DBLet(Rs!impornet, "N")), Mens, CStr(DBLet(Rs!cantfact, "N")))

        NumF = NumF + 1
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing

    InsertarLineasCalibreFactu = True
    
EInsertarLin:
    If Err.Number <> 0 Or Not b Then
        InsertarLineasCalibreFactu = False
        If b Then Mens = ""
        MuestraError Err.Number, "Insertar lineas de calibre de factura.", Err.Description & " " & Mens
    Else
        InsertarLineasCalibreFactu = True
    End If
End Function



