VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFacturaTra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'COMPRAS
'FACTURA: se corresponde con la tabla ariges.scafpc (cabecera de facturas de Proveedores)
'-------------------------------------------------------------------------

'Variables locales que contienen valores de propiedad
Private mCodTrans As String 'Cod. agencia de transporte
Private mNumFactu As String 'Nº Factura
Private mFecFactu As String 'FEcha Factura
Private mFecRecep As String 'FEcha Recepcion

Private mCodmacta As String 'cuenta de la agencia de transporte para insertar in spagop

Private mNomTrans As String 'nombre del proveedor
Private mDomTrans As String 'domicilio
Private mCodpobla As String 'cpostal
Private mPobTrans As String 'poblacion
Private mProTrans As String 'provincia
Private mNifTrans As String 'NIF agencia de transporte
Private mTelTrans As String 'telefono

Private mCodForpa As Integer 'Cod. Forma de pago
Private mCodTraba As String 'Cod. trabajador factura

Private mBrutofac As Currency 'total importe bruto de las lineas de la factura(sin dtos ni IVA)
Private mDtoPpago As Currency 'descuento pronto pago
Private mDtoGnral As Currency 'descuento general
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

Private mPorcRet As Currency    '% de retencion
Private mImpReten As Currency 'Importe al aplicar a la base imponible la retencion


Private mPorcrec1 As Currency   '% de recargo 1
Private mPorcrec2 As Currency   '% de recargo 2
Private mPorcrec3 As Currency   '% de recargo 3

Private mImpoiva1 As Currency   'Importe de iva 1
Private mImpoiva2 As Currency   'Importe de iva 2
Private mImpoiva3 As Currency   'Importe de iva 3
Private mImporec1 As Currency   'Importe de recargo 1
Private mImporec2 As Currency   'Importe de recargo 2
Private mImporec3 As Currency   'Importe de recargo 3

Private mTotalfac As Currency  'total factura (con dtos y con IVA)

Private mBanPr As String 'Banco propio que ponemos como cuenta prevista de pago
Private mCtaPrev As String 'Cuenta prevista de pago


Private mtipo As Byte   '0 = TRANSPORTE  1 = COMISION


'DAVID###
Private mCCC_Entidad As String
Private mCCC_Oficina As String
Private mCCC_CC As String
Private mCCC_CTa As String

'[Monica]22/11/2013: Tema iban
Private mCCC_Iban As String


Dim vvIban As String

'------------------------------------------------
'Propiedades del modulo CFacturaCom
'------------------------------------------------

'**** Proveedor de la Factura

Public Property Let Transportista(ByVal vData As String)
     mCodTrans = vData
End Property

Public Property Get Transportista() As String
     Transportista = mCodTrans
End Property


'**** Nº de la Factura

Public Property Let NumFactu(ByVal vData As String)
     mNumFactu = vData
End Property

Public Property Get NumFactu() As String
     NumFactu = mNumFactu
End Property


'**** Fecha de la Factura

Public Property Let FecFactu(ByVal vData As String)
     mFecFactu = vData
End Property

Public Property Get FecFactu() As String
     FecFactu = mFecFactu
End Property


'**** Fecha Recepcion de la Factura

Public Property Let FecRecep(ByVal vData As String)
     mFecRecep = vData
End Property

Public Property Get FecRecep() As String
     FecRecep = mFecRecep
End Property


'**** cta contable (codmacta) del transportista de la Factura

Public Property Let CtaTrans(ByVal vData As String)
     mCodmacta = vData
End Property

Public Property Get CtaTrans() As String
     CtaTrans = mCodmacta
End Property


'**** Nombre del Transportista de la Factura

Public Property Let NombreTrans(ByVal vData As String)
     mNomTrans = vData
End Property

Public Property Get NombreTrans() As String
     NombreTrans = mNomTrans
End Property

'**** Domicilio del Transportista de la Factura

Public Property Let DomicilioTrans(ByVal vData As String)
     mDomTrans = vData
End Property

Public Property Get DomicilioTrans() As String
     DomicilioTrans = mDomTrans
End Property


'**** CPostal del Transportista de la Factura

Public Property Let CPostalTrans(ByVal vData As String)
     mCodpobla = vData
End Property

Public Property Get CPostalTrans() As String
     CPostalTrans = mCodpobla
End Property


'**** Poblacion del Transpoortista de la Factura

Public Property Let PoblacionTrans(ByVal vData As String)
     mPobTrans = vData
End Property

Public Property Get PoblacionTrans() As String
     PoblacionTrans = mPobTrans
End Property

'**** Provincia del Transportista de la Factura

Public Property Let ProvinciaTrans(ByVal vData As String)
     mProTrans = vData
End Property

Public Property Get ProvinciaTrans() As String
     ProvinciaTrans = mProTrans
End Property


'**** NIF del Transportista de la Factura

Public Property Let NIFTrans(ByVal vData As String)
     mNifTrans = vData
End Property

Public Property Get NIFTrans() As String
     NIFTrans = mNifTrans
End Property


'**** Telefono del Transportista de la Factura

Public Property Let TelefTrans(ByVal vData As String)
     mTelTrans = vData
End Property

Public Property Get TelefTrans() As String
     TelefTrans = mTelTrans
End Property


'**** Trabajador de la Factura

Public Property Let Trabajador(ByVal vData As String)
     mCodTraba = vData
End Property

Public Property Get Trabajador() As String
     Trabajador = mCodTraba
End Property


'**** Forma Pago de la Factura

Public Property Let ForPago(ByVal vData As Integer)
     mCodForpa = vData
End Property

Public Property Get ForPago() As Integer
     ForPago = mCodForpa
End Property


'**** Importe Bruto Factura

Public Property Let BrutoFac(ByVal vData As Currency)
     mBrutofac = vData
End Property

Public Property Get BrutoFac() As Currency
     BrutoFac = mBrutofac
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

'**** Importe pronto pago de la Factura

Public Property Let ImpPPago(ByVal vData As Currency)
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

Public Property Let TotalFac(ByVal vData As Currency)
     mTotalfac = vData
End Property

Public Property Get TotalFac() As Currency
     TotalFac = mTotalfac
End Property


'**** Banco propio de pago de la Factura

Public Property Let BancoPr(ByVal vData As String)
     mBanPr = vData
End Property

Public Property Get BancoPr() As String
     BancoPr = mBanPr
End Property

'**** Cuenta Prevista de pago de la Factura
Public Property Let CuentaPrev(ByVal vData As String)
     mCtaPrev = vData
End Property
Public Property Get CuentaPrev() As String
     CuentaPrev = mCtaPrev
End Property


'**** Cuenta banacaria: ENTIDAD
Public Property Let CCC_Entidad(ByVal vData As String)
     mCCC_Entidad = vData
End Property
Public Property Get CCC_Entidad() As String
     CCC_Entidad = mCCC_Entidad
End Property

'**** Cuenta banacaria: OFICINA
Public Property Let CCC_Oficina(ByVal vData As String)
     mCCC_Oficina = vData
End Property
Public Property Get CCC_Oficina() As String
     CCC_Oficina = mCCC_Oficina
End Property

'**** Cuenta banacaria: Codigo control
Public Property Let CCC_CC(ByVal vData As String)
     mCCC_CC = vData
End Property
Public Property Get CCC_CC() As String
     CCC_CC = mCCC_CC
End Property

'Cuenta banacaria.  Cuenta
Public Property Let CCC_CTa(ByVal vData As String)
     mCCC_CTa = vData
End Property
Public Property Get CCC_CTa() As String
     CCC_CTa = mCCC_CTa
End Property

'[Monica]22/11/2013: Tema Iban
'**** Cuenta banacaria: Iban
Public Property Let CCC_Iban(ByVal vData As String)
     mCCC_Iban = vData
End Property
Public Property Get CCC_Iban() As String
     CCC_Iban = mCCC_Iban
End Property


'**** % de retencion de la Factura

Public Property Let PorcReten(ByVal vData As Currency)
     mPorcRet = vData
End Property

Public Property Get PorcReten() As Currency
     PorcReten = mPorcRet
End Property

'**** Importe retencion de la Factura

Public Property Let ImpReten(ByVal vData As Currency)
     mImpReten = vData
End Property

Public Property Get ImpReten() As Currency
     ImpReten = mImpReten
End Property




Public Property Let EsComisionista(ByVal vData As Byte)
     mtipo = vData
End Property

Public Property Get EsComisionista() As Byte
     EsComisionista = mtipo
End Property


'------------------------------------------------
'Procedimientos del modulo CFacturaTra
'------------------------------------------------

Public Function LeerDatos(vCodTrans As String, vNumfactu As String, vFecfactu As String) As Boolean
'Leer los datos de una factura almacenada en la tabla scafpc
'Lee de la BD: Ariges, Tabla: scafpc
'OUT: True si lee los datos correctamente
Dim Rs As ADODB.Recordset
Dim Mens As String
Dim Sql As String

    On Error GoTo ELeer
    
    LeerDatos = False
    Sql = "SELECT codtrans,numfactu,fecfactu,codforpa,brutofac,totalfac,nomtrans,domtrans,codpobla,pobtrans,protrans,niftrans FROM tcafpc "
    Sql = Sql & " WHERE codtrans=" & vCodTrans
    Sql = Sql & " AND numfactu='" & vNumfactu & "'"
    Sql = Sql & " AND fecfactu='" & Format(vFecfactu, FormatoFecha) & "'"

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    If Rs.EOF Then
        LeerDatos = False
    Else
        Mens = "Datos Factura"
        mCodTrans = Rs!codTrans
        mNumFactu = CStr(Rs!NumFactu)
        mFecFactu = CStr(Rs!FecFactu)
        mCodForpa = CStr(Rs!Codforpa)
        mBrutofac = CStr(Rs!BrutoFac)
        mTotalfac = CStr(Rs!TotalFac)
        
        mNomTrans = DBLet(Rs!NomTrans, "T")
        mDomTrans = DBLet(Rs!domtrans, "T")
        mCodpobla = DBLet(Rs!codpobla, "T")
        mPobTrans = DBLet(Rs!pobtrans, "T")
        mProTrans = DBLet(Rs!protrans, "T")
        mNifTrans = DBLet(Rs!NIFTrans, "T")
        
        
        'obterner la ctaCble del proveedor
        CtaTrans = DevuelveDesdeBDNew(cAgro, "agencias", "codmacta", "codtrans", vCodTrans, "N")
        
        
        'si venimos de proceso de insertar tenemos la cuenta prevista de pago que
        'se introdujo en el formulario de pasar albaranes a factura
        'si no venimos de moficicar y la leemos de spagos antes de actualizar
        If CuentaPrev = "" Then
            'leer la cuenta prevista de pagos de la tabla de contabilidad: conta.spagos
            'cuando vamos a borrar una factura porque modificarmos
            If vParamAplic.ContabilidadNueva Then
                Sql = DevuelveDesdeBDNew(cConta, "pagos", "ctabanc1", "codmacta", CtaTrans, "T", , "numfactu", vNumfactu, "T", "fecfactu", vFecfactu, "F")
            Else
                Sql = DevuelveDesdeBDNew(cConta, "spagop", "ctabanc1", "ctaprove", CtaTrans, "T", , "numfactu", vNumfactu, "T", "fecfactu", vFecfactu, "F")
            End If
            If Sql <> "" Then
                CuentaPrev = Sql
                LeerDatos = True
            Else
                LeerDatos = False
                Mens = "La cuenta prevista de cobro para la factura no puede ser nula."
            End If
        Else
            LeerDatos = True
        End If
        
'        LeerDatos = True
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    If Not Rs Is Nothing Then Set Rs = Nothing
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
Dim Sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single
Dim I As Byte

    On Error GoTo EInsertarTesoreria

'    b = False
    InsertarEnTesoreria = False
    CadValues2 = ""

    'Obtener el Nº de Vencimientos de la forma de pago
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
               CadValuesAux2 = "('" & CtaTrans & "', " & DBSet(NumFactu, "T") & ", '" & Format(FecFactu, FormatoFecha) & "', "
              
              'Primer Vencimiento
              '------------------------------------------------------------
              I = 1
              'FECHA VTO
              FecVenci = CDate(FecFactu)
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
             
              CadValues2 = CadValuesAux2 & I
              CadValues2 = CadValues2 & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                
              'IMPORTE del Vencimiento
              If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFac
              Else
                    ImpVenci = Round(TotalFac / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFac Then
                        ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * rsVenci!numerove), 2)
                    End If
              End If
              CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CuentaPrev, "T") & ","
              
              If vParamAplic.ContabilidadNueva Then
                    vvIban = MiFormat(mCCC_Iban, "") & MiFormat(mCCC_Entidad, "0000") & MiFormat(mCCC_Oficina, "0000") & MiFormat(mCCC_CC, "00") & MiFormat(mCCC_CTa, "0000000000")
                    
                    CadValues2 = CadValues2 & DBSet(vvIban, "T") & ","
              Else
                    'David. Para que ponga la cuenta bancaria (SI LA tiene)
                    CadValues2 = CadValues2 & DBSet(mCCC_Entidad, "T", "S") & "," & DBSet(mCCC_Oficina, "T", "S") & ","
                    CadValues2 = CadValues2 & DBSet(mCCC_CC, "T", "S") & "," & DBSet(mCCC_CTa, "T", "S") & ","
                    
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                          CadValues2 = CadValues2 & DBSet(mCCC_Iban, "T", "S") & ","
                    End If
              End If

              'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
              Sql = "Factura num.: " & NumFactu & "-" & Format(FecFactu, "dd/mm/yyyy")
              CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
              Sql = "Vto a fecha: " & Format(FecVenci1, "dd/mm/yyyy")
              CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'"
              If vParamAplic.ContabilidadNueva Then
                  CadValues2 = CadValues2 & "," & DBSet(mNomTrans, "T") & "," & DBSet(mDomTrans, "T") & "," & DBSet(mPobTrans, "T") & "," & DBSet(mCodpobla, "T") & "," & DBSet(mProTrans, "T") & "," & DBSet(mNifTrans, "T") & "'ES')"
              Else
                  CadValues2 = CadValues2 & ")"
              End If

              'Resto Vencimientos
              '--------------------------------------------------------------------
              For I = 2 To rsVenci!numerove
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

                  CadValues2 = CadValues2 & ", " & CadValuesAux2 & I & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "

                  'IMPORTE Resto de Vendimientos
                  ImpVenci = Round(TotalFac / rsVenci!numerove, 2)

                  CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", '" & CuentaPrev & "',"
                  
                  If vParamAplic.ContabilidadNueva Then
                      vvIban = MiFormat(mCCC_Iban, "") & MiFormat(mCCC_Entidad, "0000") & MiFormat(mCCC_Oficina, "0000") & MiFormat(mCCC_CC, "00") & MiFormat(mCCC_CTa, "0000000000")
                      
                      CadValues2 = CadValues2 & DBSet(vvIban, "T")
                  Else
                  
                      'David. Para que ponga la cuenta bancaria (SI LA tiene)
                      CadValues2 = CadValues2 & DBSet(mCCC_Entidad, "T", "S") & "," & DBSet(mCCC_Oficina, "T", "S") & ","
                      CadValues2 = CadValues2 & DBSet(mCCC_CC, "T", "S") & "," & DBSet(mCCC_CTa, "T", "S") & ","
                      '[Monica]22/11/2013: Tema Iban
                      If vEmpresa.HayNorma19_34Nueva = 1 Then
                          CadValues2 = CadValues2 & DBSet(mCCC_Iban, "T", "S") & ","
                      End If
                  End If
                  Sql = "Factura num.: " & NumFactu & "-" & Format(FecFactu, "dd/mm/yyyy")
                  CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
                  Sql = "Vto a fecha: " & Format(FecVenci1, "dd/mm/yyyy")
                  CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'"
                  If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & "," & DBSet(mNomTrans, "T") & "," & DBSet(mDomTrans, "T") & "," & DBSet(mPobTrans, "T") & "," & DBSet(mCodpobla, "T") & "," & DBSet(mProTrans, "T") & "," & DBSet(mNifTrans, "T") & ",'ES')"
                  Else
                        CadValues2 = CadValues2 & ")"
                  End If
              Next I
            End If
            Rs.Close
            Set Rs = Nothing
        End If
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
            cadValuesAux = "tipoforp"
            CadValuesAux2 = DevuelveDesdeBDNew(cAgro, "forpago", "nomforpa", "codforpa", ForPago, "N", cadValuesAux)
            If vParamAplic.ContabilidadNueva Then
                'insertamos e sforpa de la CONTA
                Sql = "INSERT INTO formapago(codforpa,nomforpa,tipforpa)"
                Sql = Sql & " VALUES(" & ForPago & ", " & DBSet(CadValuesAux2, "T") & ", " & cadValuesAux & ")"
            Else
                'insertamos e sforpa de la CONTA
                Sql = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
                Sql = Sql & " VALUES(" & ForPago & ", " & DBSet(CadValuesAux2, "T") & ", " & cadValuesAux & ")"
            End If
            ConnConta.Execute Sql
        End If


        
        'Insertamos en la tabla spagop de la CONTA
        'SQL = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1) "
        'David. Cuenta bancaria y descripcion textos
        If vParamAplic.ContabilidadNueva Then
            Sql = "INSERT INTO pagos (codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,iban,"
            Sql = Sql & "text1csb,text2csb,nomprove,domprove,pobprove,cpprove,proprove,nifprove,codpais) "
        
        
        Else
            Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,"
            
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
    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        b = False
        MenError = "Error al insertar en Tesoreria: " & Err.Description
    End If
    InsertarEnTesoreria = b
End Function



Public Function TraspasoAlbaranesAFactura(cadSQL As String, PorlineaAlbaran As Boolean) As Boolean
'IN -> cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'Desde Albaranes Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String
Dim Sql As String

    On Error GoTo ETraspasoAlbFac
    TraspasoAlbaranesAFactura = False
    
    
    '[Monica] 11/01/2010 : necesito una temporal para cargar las lineas de factura pq ahora acumulo costes
    Sql = "DROP TABLE IF EXISTS tmpLineas; "
    conn.Execute Sql
    
    Sql = "CREATE  TABLE `tmpLineas` (" 'TEMPORARY
    Sql = Sql & "`numalbar` int(6), "
    Sql = Sql & "`numlinea` smallint(2), "
    Sql = Sql & "`importel` decimal(10,2) "
    Sql = Sql & " ) ENGINE=InnoDB DEFAULT CHARSET=latin1"

    conn.Execute Sql
    '[Monica]
    
    conn.BeginTrans
    If vParamAplic.NumeroConta <> 0 Then ConnConta.BeginTrans
    
    
'    cadSQL = Replace(cadSQL, "slialp", "scaalp")
    
    'Cuenta Prevista de Cobro de las Facturas
    CuentaPrev = DevuelveDesdeBDNew(cAgro, "banpropi", "codmacta", "codbanpr", BancoPr, "N")
    
    If mtipo = 0 Then
        If PorlineaAlbaran Then
            b = InsertarCostesPortesLin(cadSQL, MenError)
        Else
            b = InsertarCostesPortes(cadSQL, MenError)
        End If
        
    Else
        If PorlineaAlbaran Then
            b = InsertarCostesComisionLin(cadSQL, MenError)
        Else
            b = InsertarCostesComision(cadSQL, MenError)
        End If
    
    End If
    
    'Insertar la Factura
    If b Then
        b = InsertarFactura(cadSQL, MenError)
    
        If b Then
            'Insertar en Tesoreria
            '#### Laura 15/11/2006 Recuperar facturas ALZIRA
            If frmVtasRecFactTrans.Check1(0).Value = 1 Then
            '####
                If vParamAplic.NumeroConta <> 0 Then b = InsertarEnTesoreria(MenError)
            End If
            
            
            '#### Laura 15/11/2006 Recuperar facturas ALZIRA
            If frmVtasRecFactTrans.Check1(1).Value = 1 Then
                If b Then
                    If vParamAplic.NumeroConta <> 0 Then b = Me.MarcarContabilizada(MenError)
                End If
            End If
            '####
            
            
        End If
    End If
    If b Then
        conn.CommitTrans
        If vParamAplic.NumeroConta <> 0 Then ConnConta.CommitTrans
        TraspasoAlbaranesAFactura = True
        Select Case mtipo
            Case 0
                MsgBox "La Factura de Transporte de los Albaranes seleccionados se generó correctamente.", vbInformation
            Case 1
                MsgBox "La Factura de Comisión de los Albaranes seleccionados se generó correctamente.", vbInformation
        End Select
    Else
        conn.RollbackTrans
        If vParamAplic.NumeroConta <> 0 Then ConnConta.RollbackTrans
        TraspasoAlbaranesAFactura = False
        MsgBox "ATENCIÓN:" & vbCrLf & "La Factura NO se generó correctamente!!!." & vbCrLf & MenError, vbInformation
    End If

    Exit Function


ETraspasoAlbFac:
    If Err.Number <> 0 Then
        conn.RollbackTrans
        If vParamAplic.NumeroConta <> 0 Then ConnConta.RollbackTrans
        TraspasoAlbaranesAFactura = False
        MuestraError Err.Number, "Pasando Albaranes a Factura.(Recepción)", Err.Description
    End If
End Function


Public Function TraspasoPortesVueltaAFactura(cadSQL As String) As Boolean
'IN -> cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'Desde Albaranes Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETraspasoAlbFac
    TraspasoPortesVueltaAFactura = False
    
    conn.BeginTrans
    If vParamAplic.NumeroConta <> 0 Then ConnConta.BeginTrans
    
    
'    cadSQL = Replace(cadSQL, "slialp", "scaalp")
    
    'Cuenta Prevista de Cobro de las Facturas
    CuentaPrev = DevuelveDesdeBDNew(cAgro, "banpropi", "codmacta", "codbanpr", BancoPr, "N")
    
    'Insertar la Factura
    b = InsertarFacturaPortes(cadSQL, MenError)

    If b Then
        'Insertar en Tesoreria
        '#### Laura 15/11/2006 Recuperar facturas ALZIRA
        If frmVtasRecFactTrans.Check1(0).Value = 1 Then
        '####
            If vParamAplic.NumeroConta <> 0 Then b = InsertarEnTesoreria(MenError)
        End If
        
        
        '#### Laura 15/11/2006 Recuperar facturas ALZIRA
        If frmVtasRecFactTrans.Check1(1).Value = 1 Then
            If b Then
                If vParamAplic.NumeroConta <> 0 Then b = Me.MarcarContabilizada(MenError)
            End If
        End If
        '####
        
        
    End If
    If b Then
        conn.CommitTrans
        If vParamAplic.NumeroConta <> 0 Then ConnConta.CommitTrans
        TraspasoPortesVueltaAFactura = True
        MsgBox "La Factura de Transporte de los Portes Vuelta se generó correctamente.", vbInformation
    Else
        conn.RollbackTrans
        If vParamAplic.NumeroConta <> 0 Then ConnConta.RollbackTrans
        TraspasoPortesVueltaAFactura = False
        MsgBox "ATENCIÓN:" & vbCrLf & "La Factura NO se generó correctamente!!!." & vbCrLf & MenError, vbInformation
    End If

    Exit Function


ETraspasoAlbFac:
    If Err.Number <> 0 Then
        conn.RollbackTrans
        If vParamAplic.NumeroConta <> 0 Then ConnConta.RollbackTrans
        TraspasoPortesVueltaAFactura = False
        MuestraError Err.Number, "Pasando Portes Vuelta a Factura.(Recepción)", Err.Description
    End If
End Function




Private Function InsertarCostesPortes(cadSQL As String, Optional cadErr As String) As Boolean
'Insertar las lineas de albaran_costes para cada variedad prorrateada con los kilos
Dim Sql As String, Sql2 As String, Sql3 As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim NetoTotal As Currency, RestoPortes As Currency, PortesLinea As Currency, PaletsLinea As Currency
Dim Linea As Currency

    On Error GoTo EInsertarAlb

    Sql = "select numalbar, portespag, paletspag from albaran where " & cadSQL
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        If vParamAplic.PortesKiloCaja = 0 Then ' caso de que se calculen por kilos
            Sql2 = "select sum(pesoneto) from albaran_variedad where numalbar = " & DBLet(Rs!numalbar, "N")
            Set Rs1 = New ADODB.Recordset
            Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            NetoTotal = 0
            If Not Rs1.EOF Then NetoTotal = DBLet(Rs1.Fields(0).Value, "N")
            Set Rs1 = Nothing
            
            If NetoTotal <> 0 Then
                RestoPortes = DBLet(Rs!portespag, "N")
                
                Sql2 = "select numlinea, pesoneto from albaran_variedad where numalbar = " & DBLet(Rs!numalbar, "N")
                Sql2 = Sql2 & " order by numlinea "
                
                Set Rs1 = New ADODB.Recordset
                Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not Rs1.EOF
                    PortesLinea = Round2(DBLet(Rs1!Pesoneto, "N") * DBLet(Rs!portespag, "N") / NetoTotal, 2)
                    RestoPortes = RestoPortes - PortesLinea
                    
                    Linea = DBLet(Rs1!numlinea, "N")
                    
                    Sql3 = "select count(*) from albaran_costes where numalbar = " & DBLet(Rs!numalbar, "N")
                    Sql3 = Sql3 & " and numlinea = " & DBSet(Rs1!numlinea, "N")
                    Sql3 = Sql3 & " and tipogasto = 2"
                    
                    If TotalRegistros(Sql3) = 0 Then
                        Sql = "INSERT INTO albaran_costes(numalbar, numlinea, tipogasto, codcoste, impcoste, importes, unidades, codartic) "
                        Sql = Sql & " VALUES (" & DBLet(Rs!numalbar, "N") & "," & DBSet(Rs1!numlinea, "N") & ",2," & ValorNulo & ","
                        Sql = Sql & DBSet(PortesLinea, "N") & "," & DBSet(PortesLinea, "N") & ",1," & ValorNulo & ")"
                    
                    Else
                        Sql = "update albaran_costes set impcoste = impcoste + " & DBSet(PortesLinea, "N")
                        Sql = Sql & ", importes = importes + " & DBSet(PortesLinea, "N")
                        Sql = Sql & " where numalbar = " & DBSet(Rs!numalbar, "N")
                        Sql = Sql & " and numlinea = " & DBSet(Rs1!numlinea, "N")
                        Sql = Sql & " and tipogasto = 2"
                    
                    End If
                    
                    conn.Execute Sql
                
                
                    Sql = "insert into tmpLineas (numalbar, numlinea, importel) values ("
                    Sql = Sql & DBSet(Rs!numalbar, "N") & "," & DBSet(Rs1!numlinea, "N") & ","
                    Sql = Sql & DBSet(PortesLinea, "N") & ")"
                
                    conn.Execute Sql
                
                    Rs1.MoveNext
                Wend
                Set Rs1 = Nothing
                
                If RestoPortes <> 0 Then
                    Sql = "update albaran_costes set impcoste = impcoste + " & DBSet(RestoPortes, "N")
                    Sql = Sql & ", importes = importes + " & DBSet(RestoPortes, "N")
                    Sql = Sql & " where numalbar = " & DBLet(Rs!numalbar, "N") & " and numlinea = " & DBSet(Linea, "N")
                    Sql = Sql & " and tipogasto = 2 "
                    
                    conn.Execute Sql
                    
                    Sql = "update tmpLineas set importel = importel + " & DBSet(RestoPortes, "N")
                    Sql = Sql & " where numalbar = " & DBLet(Rs!numalbar, "N") & " and numlinea = " & DBSet(Linea, "N")
                    
                    conn.Execute Sql
                    
                    
                End If
            End If
        Else
            ' caso de que se calculen por cajas
            Sql2 = "select albaran_variedad.numlinea, albaran_variedad.numcajas, albaran_variedad.codforfait, forfaits.cajaspalet from albaran_variedad, forfaits where numalbar = " & DBLet(Rs!numalbar, "N")
            Sql2 = Sql2 & " and albaran_variedad.codforfait = forfaits.codforfait "
        
            Set Rs1 = New ADODB.Recordset
            Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
            RestoPortes = DBLet(Rs!portespag, "N")
            
            While Not Rs1.EOF
                PortesLinea = 0
                Linea = DBLet(Rs1.Fields(0).Value, "N")
                
                If DBLet(Rs1.Fields(3).Value, "N") <> 0 Then
                    PaletsLinea = Round2(DBLet(Rs1.Fields(1).Value, "N") / DBLet(Rs1.Fields(3).Value, "N"), 2)
                    
                    If DBLet(Rs!paletspag, "N") <> 0 Then
                        PortesLinea = Round2(DBLet(Rs!portespag, "N") * PaletsLinea / DBLet(Rs!paletspag, "N"), 4)
                    End If
                    ' --monica:cambiado por lo de arriba pq ahora tenemos en cuenta el numero de palets de la linea de albaran.
                    'PortesLinea = Round2(DBLet(Rs!portespag, "N") / DBLet(Rs1.Fields(3).Value, "N") * DBLet(Rs1.Fields(1).Value, "N"), 4)
                    RestoPortes = RestoPortes - PortesLinea
                End If
                
                Sql3 = "select count(*) from albaran_costes where numalbar = " & DBLet(Rs!numalbar, "N")
                Sql3 = Sql3 & " and numlinea = " & DBSet(Rs1!numlinea, "N")
                Sql3 = Sql3 & " and tipogasto = 2"
                
                If TotalRegistros(Sql3) = 0 Then
                    Sql = "INSERT INTO albaran_costes(numalbar, numlinea, tipogasto, codcoste, impcoste, importes, unidades, codartic) "
                    Sql = Sql & " VALUES (" & DBLet(Rs!numalbar, "N") & "," & DBSet(Rs1!numlinea, "N") & ",2," & ValorNulo & ","
                    Sql = Sql & DBSet(PortesLinea, "N") & "," & DBSet(PortesLinea, "N") & ",1," & ValorNulo & ")"
                Else
                    Sql = "update albaran_costes set impcoste = impcoste + " & DBSet(PortesLinea, "N")
                    Sql = Sql & ", importes = importes + " & DBSet(PortesLinea, "N")
                    Sql = Sql & " where numalbar = " & DBSet(Rs!numalbar, "N")
                    Sql = Sql & " and numlinea = " & DBSet(Rs1!numlinea, "N")
                    Sql = Sql & " and tipogasto = 2"
                End If
                conn.Execute Sql
                
                Sql = "insert into tmpLineas (numalbar, numlinea, importel) values ("
                Sql = Sql & DBSet(Rs!numalbar, "N") & "," & DBSet(Rs1!numlinea, "N") & ","
                Sql = Sql & DBSet(PortesLinea, "N") & ")"
            
                conn.Execute Sql
                
                Rs1.MoveNext
            Wend
            
            Set Rs1 = Nothing
                
            If RestoPortes <> 0 Then
                Sql = "update albaran_costes set impcoste = impcoste + " & DBSet(RestoPortes, "N")
                Sql = Sql & ", importes = importes + " & DBSet(RestoPortes, "N")
                Sql = Sql & " where numalbar = " & DBLet(Rs!numalbar, "N") & " and numlinea = " & DBSet(Linea, "N")
                Sql = Sql & " and tipogasto = 2 "
                
                conn.Execute Sql
                
                Sql = "update tmpLineas set importel = importel + " & DBSet(RestoPortes, "N")
                Sql = Sql & " where numalbar = " & DBLet(Rs!numalbar, "N") & " and numlinea = " & DBSet(Linea, "N")
                
                conn.Execute Sql
            
            End If
        End If ' del caso en el que los calculos de portes se hagan por cajas
        Rs.MoveNext
    Wend
    
    InsertarCostesPortes = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCostesPortes = False
        cadErr = "Error al insertar líneas de Portes de albarán: " & Err.Description
    Else
        InsertarCostesPortes = True
    End If
End Function


Private Function InsertarCostesPortesLin(cadSQL As String, Optional cadErr As String) As Boolean
'Insertar las lineas de albaran_costes para cada variedad prorrateada con los kilos
Dim Sql As String, Sql2 As String, Sql3 As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim NetoTotal As Currency, RestoPortes As Currency, PortesLinea As Currency, PaletsLinea As Currency
Dim Linea As Currency

    On Error GoTo EInsertarAlb


    Sql = "select numalbar, numlinea, portespag from tmpalbaranes where codusu = " & vUsu.codigo
    Sql = Sql & " and portespag <> 0 "
    Sql = Sql & " order by 1,2"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql3 = "select count(*) from albaran_costes where numalbar = " & DBLet(Rs!numalbar, "N")
        Sql3 = Sql3 & " and numlinea = " & DBSet(Rs!numlinea, "N")
        Sql3 = Sql3 & " and tipogasto = 2"
        
        If TotalRegistros(Sql3) = 0 Then
            Sql = "INSERT INTO albaran_costes(numalbar, numlinea, tipogasto, codcoste, impcoste, importes, unidades, codartic) "
            Sql = Sql & " VALUES (" & DBLet(Rs!numalbar, "N") & "," & DBSet(Rs!numlinea, "N") & ",2," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!portespag, "N") & "," & DBSet(Rs!portespag, "N") & ",1," & ValorNulo & ")"
        
        Else
            Sql = "update albaran_costes set impcoste = impcoste + " & DBSet(Rs!portespag, "N")
            Sql = Sql & ", importes = importes + " & DBSet(Rs!portespag, "N")
            Sql = Sql & " where numalbar = " & DBSet(Rs!numalbar, "N")
            Sql = Sql & " and numlinea = " & DBSet(Rs!numlinea, "N")
            Sql = Sql & " and tipogasto = 2"
        
        End If
        
        conn.Execute Sql
    
        Sql = "insert into tmpLineas (numalbar, numlinea, importel) values ("
        Sql = Sql & DBSet(Rs!numalbar, "N") & "," & DBSet(Rs!numlinea, "N") & ","
        Sql = Sql & DBSet(Rs!portespag, "N") & ")"
    
        conn.Execute Sql
                
        Rs.MoveNext
    Wend

    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCostesPortesLin = False
        cadErr = "Error al insertar líneas de Portes Linea de albarán: " & Err.Description
    Else
        InsertarCostesPortesLin = True
    End If
End Function







Private Function InsertarFactura(cadSQL As String, cadError As String) As Boolean
'Insertamos en las tablas de factura de tranportista: tcafpc, tcafpa, tlifpc
Dim b As Boolean
    
    InsertarFactura = False
    
    
    'Cabecera Factura
    cadError = ""
    b = InsertarCabeceraFactu(cadError)
    If Not b Then
'        cadError = "Error al insertar cabecera de factura"
        Exit Function
    End If
    
    'Cabeceras Albaranes de la Factura
    b = InsertarCabAlbaranesFactu(cadSQL, cadError)
    If Not b Then
'        cadError = "Error al insertar cabecera albaranes de factura"
        Exit Function
    End If
    
    'Insertar lineas de la factura
    b = InsertarLineasFactu(cadSQL, cadError)
    If Not b Then
'        cadError = "Error al insertar líneas de factura"
        Exit Function
    End If
    
    'Insertar cabeceras de portes de vuelta
    b = InsertarCabPortesVuelta(cadSQL, cadError)
    If Not b Then
'        cadError = "Error al insertar líneas de factura"
        Exit Function
    End If
    
    
    
    InsertarFactura = True
End Function

Private Function InsertarFacturaPortes(cadSQL As String, cadError As String) As Boolean
'Insertamos en las tablas de factura de tranportista: tcafpc, tcafpa, tlifpc
Dim b As Boolean
    
    InsertarFacturaPortes = False
    
    
    'Cabecera Factura
    cadError = ""
    b = InsertarCabeceraFactu(cadError)
    If Not b Then
'        cadError = "Error al insertar cabecera de factura"
        Exit Function
    End If
    
    'Insertar cabeceras de portes de vuelta
    b = InsertarCabPortesVuelta(cadSQL, cadError)
    If Not b Then
'        cadError = "Error al insertar líneas de factura"
        Exit Function
    End If
    
    
    InsertarFacturaPortes = True
End Function


Private Function InsertarCabeceraFactu(Optional cadErr As String) As Boolean
'Inserta la cabecera de la factura en la tabla: scafpc
Dim Sql As String
Dim vTrans As CTransportista
Dim Nulo2 As String
Dim Nulo3 As String
Dim Nulo4 As String
Dim Nulo5 As String

    On Error GoTo EInsertar

    'Obtener los datos del Transportista: nombre, domicilio, NIF,....
    Set vTrans = New CTransportista
    If vTrans.Existe(Transportista) Then
        vTrans.LeerDatos (Transportista)
        '#### Leer estos datos de la tabla scaalpr y no de sprove
        NombreTrans = vTrans.Nombre
        DomicilioTrans = vTrans.Domicilio
        CPostalTrans = vTrans.CPostal
        PoblacionTrans = vTrans.Poblacion
        ProvinciaTrans = vTrans.Provincia
        NIFTrans = vTrans.NIF
        TelefTrans = vTrans.TfnoAdmon
        CtaTrans = vTrans.CuentaCble
    End If

    'Insertar en la tabla cabecera de la factura de compras
    Sql = "INSERT INTO tcafpc (codtrans,numfactu,fecfactu,fecrecep,nomtrans,domtrans,codpobla,pobtrans,protrans,niftrans,teltrans,codforpa,brutofac,dtoppago,dtognral,impppago,impgnral,baseiva1,baseiva2,baseiva3, "
    Sql = Sql & "tipoiva1,tipoiva2,tipoiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,retfacpr,trefacpr,totalfac,intconta, tipo) "
    Sql = Sql & " VALUES (" & Transportista & "," & DBSet(NumFactu, "T") & "," & DBSet(FecFactu, "F") & "," & DBSet(FecRecep, "F") & ","
    Sql = Sql & DBSet(NombreTrans, "T", "N") & "," & DBSet(DomicilioTrans, "T", "N") & "," & DBSet(CPostalTrans, "T", "N") & "," & DBSet(PoblacionTrans, "T", "N") & "," & DBSet(ProvinciaTrans, "T", "N") & ","
    Sql = Sql & DBSet(NIFTrans, "T", "N") & "," & DBSet(TelefTrans, "T", "N") & "," & ForPago & ","
    Sql = Sql & DBSet(BrutoFac, "N", "N") & ","
    Sql = Sql & DBSet(DtoPPago, "N", "N") & "," & DBSet(DtoGnral, "N", "N") & ","
    Sql = Sql & DBSet(ImpPPago, "N", "N") & "," & DBSet(ImpGnral, "N", "N") & ","
    Sql = Sql & DBSet(BaseIVA1, "N", "N") & "," & DBSet(BaseIVA2, "N", "S") & "," & DBSet(BaseIVA3, "N", "S") & ","
    Nulo2 = "N"
    Nulo3 = "N"
    If BaseIVA2 = 0 Then Nulo2 = "S"
    If BaseIVA3 = 0 Then Nulo3 = "S"
    Sql = Sql & DBSet(TipoIVA1, "N", "N") & "," & DBSet(TipoIVA2, "N", Nulo2) & "," & DBSet(TipoIVA3, "N", Nulo3) & ","
    Sql = Sql & DBSet(PorceIVA1, "N", "N") & "," & DBSet(PorceIVA2, "N", Nulo2) & "," & DBSet(PorceIVA3, "N", Nulo3) & ","
    Sql = Sql & DBSet(ImpIVA1, "N", "N") & "," & DBSet(ImpIVA2, "N", Nulo2) & "," & DBSet(ImpIVA3, "N", Nulo3) & ","
    If PorcReten = 0 Then Nulo4 = "S"
    If ImpReten = 0 Then Nulo5 = "S"
    Sql = Sql & DBSet(PorcReten, "N", Nulo4) & ","
    Sql = Sql & DBSet(ImpReten, "N", Nulo5) & ","
    Sql = Sql & DBSet(TotalFac, "N", "N") & ",0," & DBSet(mtipo, "N") & ")"
    
    conn.Execute Sql
    Set vTrans = Nothing
    InsertarCabeceraFactu = True

EInsertar:
    If Err.Number <> 0 Then
        Set vTrans = Nothing
        InsertarCabeceraFactu = False
        cadErr = "Error al insertar cabecera factura: " & Err.Description
    Else
        InsertarCabeceraFactu = True
    End If
End Function



Private Function InsertarCabAlbaranesFactu(cadSQL As String, Optional cadErr As String) As Boolean
'Insertar las cabeceras de los Albaranes de la factura en la tabla: scafpa
Dim Sql As String

    On Error GoTo EInsertarAlb

    Sql = "INSERT INTO tcafpa (codtrans,numfactu,fecfactu,numalbar,fechaalb,importel,observa1,observa2,observa3,observa4,observa5) "
    Sql = Sql & " SELECT " & Transportista & " as codtrans," & DBSet(NumFactu, "T") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu," & " numalbar,fechaalb,"
    
    Select Case mtipo
        Case 0 ' transportista
            Sql = Sql & "portespag , " & ValorNulo & ", " & ValorNulo & ", " & ValorNulo & ", " & ValorNulo & ", " & ValorNulo
        Case 1 ' comisionista
            Sql = Sql & "comisionespag , " & ValorNulo & ", " & ValorNulo & ", " & ValorNulo & ", " & ValorNulo & ", " & ValorNulo
    End Select
    
    Sql = Sql & " FROM albaran WHERE " & cadSQL
    conn.Execute Sql

    InsertarCabAlbaranesFactu = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCabAlbaranesFactu = False
        cadErr = "Error al insertar cabecera albaranes Factura: " & Err.Description
    Else
        InsertarCabAlbaranesFactu = True
    End If
End Function

Private Function InsertarCabPortesVFactu(cadSQL As String, Optional cadErr As String) As Boolean
'Insertar los portes de vuelta en la tabla: scafpv
Dim Sql As String

    On Error GoTo EInsertarAlb

    Sql = "INSERT INTO tcafpv (codtrans,numfactu,fecfactu,codmacta,importel) "
    Sql = Sql & " SELECT " & Transportista & " as codtrans," & DBSet(NumFactu, "T") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu," & " codmacta,importel "
    Sql = Sql & " FROM tmpportesvuelta WHERE " & cadSQL
    conn.Execute Sql

    InsertarCabPortesVFactu = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCabPortesVFactu = False
        cadErr = "Error al insertar Portes Vuelta Factura: " & Err.Description
    Else
        InsertarCabPortesVFactu = True
    End If
End Function



Private Function InsertarCabPortesVuelta(cadSQL As String, Optional cadErr As String) As Boolean
'Insertar las cabeceras de los Albaranes de la factura en la tabla: scafpa
Dim Sql As String

    On Error GoTo EInsertarAlb

    Sql = "INSERT INTO tcafpv (codtrans,numfactu,fecfactu,codmacta,importel) "
    Sql = Sql & " SELECT " & Transportista & " as codtrans," & DBSet(NumFactu, "T") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu," & " codmacta,importel "
    Sql = Sql & " FROM tmpportesv WHERE codusu = " & vUsu.codigo & " and codtrans = " & DBSet(Transportista, "N")
    Sql = Sql & " and numfactu = " & DBSet(NumFactu, "T")
    Sql = Sql & " and fecfactu = " & DBSet(FecFactu, "F")
    conn.Execute Sql

    InsertarCabPortesVuelta = True
    Exit Function
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCabPortesVuelta = False
        cadErr = "Error al insertar portes vuelta Factura: " & Err.Description
    Else
        InsertarCabPortesVuelta = True
    End If
End Function



Private Function InsertarLineasFactu(cadSQL As String, cadErr As String) As Boolean
'Insertar las lineas de los Albaranes de la factura en la tabla: slifpc
Dim Sql As String

    On Error GoTo EInsertarLin

    Sql = "INSERT INTO tlifpc (codtrans,numfactu,fecfactu,numalbar,numlinea,importel) "
    Sql = Sql & " SELECT " & Transportista & " as codtrans," & DBSet(NumFactu, "T") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu," & " numalbar, numlinea, importel "
'    SQL = SQL & " FROM albaran_costes, albaran WHERE " & cadSQL
    Sql = Sql & " FROM tmpLineas "
    
'    Select Case mtipo
'        Case 0 'transportista
'            SQL = SQL & " and tipogasto = 2 and albaran.numalbar = albaran_costes.numalbar "
'        Case 1 ' Comsionista
'            SQL = SQL & " and tipogasto = 3 and albaran.numalbar = albaran_costes.numalbar "
'    End Select
    conn.Execute Sql

    InsertarLineasFactu = True
    
EInsertarLin:
    If Err.Number <> 0 Then
        InsertarLineasFactu = False
        cadErr = "Error al insertar líneas de factura: " & Err.Description
    Else
        InsertarLineasFactu = True
    End If
End Function


'Private Function EliminarAlbaranes(cadSQL As String, cadError As String) As Boolean
''Eliminamos de las tablas de Albaranes: scaalp, slialp
'Dim Sql As String
'
'    On Error GoTo EEliminar
'
'    EliminarAlbaranes = False
'
'    'ELiminar lineas
'    Sql = "DELETE FROM slialp "
'    Sql = Sql & " WHERE " & Replace(cadSQL, "scaalp", "slialp")
'    Conn.Execute Sql
'
'    'Eliminar Cabeceras Albaranes
'    Sql = "DELETE FROM scaalp "
'    Sql = Sql & " WHERE " & cadSQL
'    Conn.Execute Sql
'
'    EliminarAlbaranes = True
'    Exit Function
'
'EEliminar:
'    EliminarAlbaranes = False
'    cadError = "Error al Eliminar albaranes: " & Err.Description
'End Function



Public Function CalcularDatosFactura(cadWhere As String, Transportista As String, Optional Insertada As Boolean) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim Rs As ADODB.Recordset
Dim I As Integer

Dim Sql As String
Dim cadAux As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotImpIVA As Currency

Dim ImpAux As Currency
Dim ImpIVA As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA

Dim vTrans As CTransportista

    CalcularDatosFactura = False
    On Error GoTo ECalcular

    'Obtener los datos del Transportista: nombre, domicilio, NIF,....
    Set vTrans = New CTransportista
    If vTrans.Existe(Transportista) Then
        vTrans.LeerDatos (Transportista)
        '#### Leer estos datos de la tabla scaalpr y no de sprove
        PorcReten = vTrans.PorcReten
    Else
        PorcReten = ValorNulo
    End If

    If Not Insertada Then
        If Not vTrans.EsComisionista = 1 Then
            Sql = "SELECT sum(portespag) as bruto"
        Else
            Sql = "SELECT sum(comisionespag) as bruto"
        End If
        Sql = Sql & " FROM albaran "
        Sql = Sql & " WHERE " & cadWhere
    Else
        Sql = "SELECT sum(importel) as bruto"
        Sql = Sql & " FROM tcafpa "
        Sql = Sql & " WHERE " & cadWhere
    End If
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    TotBruto = 0
    I = 1
    
'    While Not Rs.EOF
        'Aqui vamos acumulando la suma del importe bruto de las lineas
        TotBruto = TotBruto + DBLet(Rs!Bruto, "N")
        
    Rs.Close
    
    If Not Insertada Then
        Sql = "select sum(importel) as bruto"
        Sql = Sql & " from tmpportesv "
        Sql = Sql & " where codusu = " & vUsu.codigo & " and codtrans = " & DBSet(Transportista, "N")
    Else
        Sql = "select sum(importel) as bruto"
        Sql = Sql & " from tcafpv "
        Sql = Sql & " where " & Replace(cadWhere, "tcafpa", "tcafpv")
    End If
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        TotBruto = TotBruto + DBLet(Rs!Bruto, "N")
        
        
        ImpBImIVA = DBLet(TotBruto, "N")
        
        '---- Laura: 24/10/2006
        ImpAux = CalcularPorcentaje(TotBruto, DtoPPago, 2)
        ImpBImIVA = ImpBImIVA - ImpAux '(bruto - DtoPP)
        
        '---- Laura: 24/10/2006
        ImpAux = CalcularPorcentaje(TotBruto, DtoGnral, 2)
        ImpBImIVA = ImpBImIVA - ImpAux '(bruto - Dtogn)
        
        ' aplicamos la retencion
        ImpAux = CalcularPorcentaje(ImpBImIVA, PorcReten, 2)
        ImpReten = ImpAux
        
        'Obtener el % de IVA
        If vParamAplic.NumeroConta <> 0 Then
            cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(vParamAplic.CodIvaTrans), "N")
        Else
            cadAux = DevuelveDesdeBDNew(cAgro, "tiposiva", "porceiva", "codigiva", CStr(vParamAplic.CodIvaTrans), "N")
        End If
        
        
        
        
        'aplicar el IVA a la base imponible de ese tipo
'        ImpAux = CalcularDto(CStr(ImpBImIVA), cadAux)
'        '---- Laura: modificado 27/09/2006
''        ImpIVA = ImpAux
'        ImpIVA = Round(ImpAux, 2)
        '----
        '---- Laura: 24/10/2006
        If cadAux = "" Then cadAux = "0"
        ImpIVA = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotImpIVA = TotImpIVA + ImpIVA
        
        
        Select Case I
            Case 1  'IVA 1
                TipoIVA1 = vParamAplic.CodIvaTrans
                
                BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA1 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA1 = ImpIVA
                
            Case 2  'IVA 2
                TipoIVA2 = 0 'RS!CodigIVA
                
                BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA2 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA2 = ImpIVA

            Case 3  'IVA 3
                TipoIVA3 = 0 'RS!CodigIVA
                
                BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA3 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA3 = ImpIVA
        End Select
        I = I + 1
'        Rs.MoveNext
'    Wend
    
    Rs.Close

    'TOTALES
    BrutoFac = TotBruto
    
    'Aplicarle el dto ppago
'    ImpPPago = CCur(CalcularDto(CStr(TotBruto), CStr(DtoPPago)))
'    '---- Laura: 27/09/2006
'    ImpPPago = Round(ImpPPago, 2)
'    '----
    '---- Laura: 24/10/2006
    ImpPPago = CalcularPorcentaje(TotBruto, DtoPPago, 2)
    '----
    
    'Aplicarle el dto general
'    ImpGnral = CCur(CalcularDto(CStr(TotBruto), CStr(DtoGnral)))
'    '---- Laura: 27/09/2006
'    ImpGnral = Round(ImpGnral, 2)
'    '----
    '---- Laura: 24/10/2006
    ImpGnral = CalcularPorcentaje(TotBruto, DtoGnral, 2)
    '----
    
    'Base Imponible
    BaseImp = TotBruto - ImpPPago - ImpGnral
    
    ImpReten = CalcularPorcentaje(BaseImp, PorcReten, 2)
    
    'TOTAL de la factura
    TotalFac = BaseImp + TotImpIVA - ImpReten
    
    Set Rs = Nothing
    CalcularDatosFactura = True
    
ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFactura = False
    Else
        CalcularDatosFactura = True
    End If
End Function



Public Function CalcularDatosFacturaSinAlbaran(cadWhere As String, Transportista As String, Optional Insertada As Boolean) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim Rs As ADODB.Recordset
Dim I As Integer

Dim Sql As String
Dim cadAux As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotImpIVA As Currency

Dim ImpAux As Currency
Dim ImpIVA As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA

Dim vTrans As CTransportista

    CalcularDatosFacturaSinAlbaran = False
    On Error GoTo ECalcular

    'Obtener los datos del Transportista: nombre, domicilio, NIF,....
    Set vTrans = New CTransportista
    If vTrans.Existe(Transportista) Then
        vTrans.LeerDatos (Transportista)
        '#### Leer estos datos de la tabla scaalpr y no de sprove
        PorcReten = vTrans.PorcReten
    Else
        PorcReten = ValorNulo
    End If

    
    TotBruto = 0
    I = 1
    
    Sql = "select sum(importel) as bruto"
    Sql = Sql & " from tmpportesv "
    Sql = Sql & " where codusu = " & vUsu.codigo & " and codtrans = " & DBSet(Transportista, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        TotBruto = TotBruto + DBLet(Rs!Bruto, "N")
        
        
        ImpBImIVA = DBLet(TotBruto, "N")
        
        '---- Laura: 24/10/2006
        ImpAux = CalcularPorcentaje(TotBruto, DtoPPago, 2)
        ImpBImIVA = ImpBImIVA - ImpAux '(bruto - DtoPP)
        
        '---- Laura: 24/10/2006
        ImpAux = CalcularPorcentaje(TotBruto, DtoGnral, 2)
        ImpBImIVA = ImpBImIVA - ImpAux '(bruto - Dtogn)
        
        ' aplicamos la retencion
        ImpAux = CalcularPorcentaje(ImpBImIVA, PorcReten, 2)
        ImpReten = ImpAux
        
        'Obtener el % de IVA
        If vParamAplic.NumeroConta <> 0 Then
            cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(vParamAplic.CodIvaTrans), "N")
        Else
            cadAux = DevuelveDesdeBDNew(cAgro, "tiposiva", "porceiva", "codigiva", CStr(vParamAplic.CodIvaTrans), "N")
        End If
        
        
        
        
        'aplicar el IVA a la base imponible de ese tipo
'        ImpAux = CalcularDto(CStr(ImpBImIVA), cadAux)
'        '---- Laura: modificado 27/09/2006
''        ImpIVA = ImpAux
'        ImpIVA = Round(ImpAux, 2)
        '----
        '---- Laura: 24/10/2006
        If cadAux = "" Then cadAux = "0"
        ImpIVA = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotImpIVA = TotImpIVA + ImpIVA
        
        
        Select Case I
            Case 1  'IVA 1
                TipoIVA1 = vParamAplic.CodIvaTrans
                
                BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA1 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA1 = ImpIVA
                
            Case 2  'IVA 2
                TipoIVA2 = 0 'RS!CodigIVA
                
                BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA2 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA2 = ImpIVA

            Case 3  'IVA 3
                TipoIVA3 = 0 'RS!CodigIVA
                
                BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA3 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA3 = ImpIVA
        End Select
        I = I + 1
'        Rs.MoveNext
'    Wend
    
    Rs.Close

    'TOTALES
    BrutoFac = TotBruto
    
    'Aplicarle el dto ppago
'    ImpPPago = CCur(CalcularDto(CStr(TotBruto), CStr(DtoPPago)))
'    '---- Laura: 27/09/2006
'    ImpPPago = Round(ImpPPago, 2)
'    '----
    '---- Laura: 24/10/2006
    ImpPPago = CalcularPorcentaje(TotBruto, DtoPPago, 2)
    '----
    
    'Aplicarle el dto general
'    ImpGnral = CCur(CalcularDto(CStr(TotBruto), CStr(DtoGnral)))
'    '---- Laura: 27/09/2006
'    ImpGnral = Round(ImpGnral, 2)
'    '----
    '---- Laura: 24/10/2006
    ImpGnral = CalcularPorcentaje(TotBruto, DtoGnral, 2)
    '----
    
    'Base Imponible
    BaseImp = TotBruto - ImpPPago - ImpGnral
    
    ImpReten = CalcularPorcentaje(BaseImp, PorcReten, 2)
    
    'TOTAL de la factura
    TotalFac = BaseImp + TotImpIVA - ImpReten
    
    Set Rs = Nothing
    CalcularDatosFacturaSinAlbaran = True
    Exit Function

ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFacturaSinAlbaran = False
    End If
End Function






'#### Laura 15/11/2006 Recuperar facturas ALZIRA
Public Function MarcarContabilizada(Optional cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    MarcarContabilizada = False
    
    Sql = "UPDATE tcafpc SET intconta=1 "
    Sql = Sql & " WHERE codtrans=" & Me.Transportista & " AND numfactu=" & DBSet(Me.NumFactu, "T") & " AND fecfactu=" & DBSet(Me.FecFactu, "F")
    conn.Execute Sql
    
    MarcarContabilizada = True
    Exit Function
    
EActualizar:
    MarcarContabilizada = False
    cadErr = "Error al marcar como contabilizada:" & Err.Description
End Function



Private Function InsertarCostesComision(cadSQL As String, Optional cadErr As String) As Boolean
'Insertar las lineas de albaran_costes para cada variedad prorrateada con los kilos
Dim Sql As String, Sql2 As String, Sql3 As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim NetoTotal As Currency, RestoComisiones As Currency, ComisionesLinea As Currency, PaletsLinea As Currency
Dim Linea As Currency

    On Error GoTo EInsertarAlb

    Sql = "select numalbar, comisionespag from albaran where " & cadSQL
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
         ' caso de que se calculen por kilos
            Sql2 = "select sum(pesoneto) from albaran_variedad where numalbar = " & DBLet(Rs!numalbar, "N")
            Set Rs1 = New ADODB.Recordset
            Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            NetoTotal = 0
            If Not Rs1.EOF Then NetoTotal = DBLet(Rs1.Fields(0).Value, "N")
            Set Rs1 = Nothing
            
            If NetoTotal <> 0 Then
                RestoComisiones = DBLet(Rs!comisionespag, "N")
                
                Sql2 = "select numlinea, pesoneto from albaran_variedad where numalbar = " & DBLet(Rs!numalbar, "N")
                Sql2 = Sql2 & " order by numlinea "
                
                Set Rs1 = New ADODB.Recordset
                Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not Rs1.EOF
                    ComisionesLinea = Round2(DBLet(Rs1!Pesoneto, "N") * DBLet(Rs!comisionespag, "N") / NetoTotal, 2)
                    RestoComisiones = RestoComisiones - ComisionesLinea
                    
                    Linea = DBLet(Rs1!numlinea, "N")
                    
                    
                    Sql3 = "select count(*) from albaran_costes where numalbar = " & DBLet(Rs!numalbar, "N")
                    Sql3 = Sql3 & " and numlinea = " & DBSet(Rs1!numlinea, "N")
                    Sql3 = Sql3 & " and tipogasto = 3"
                    
                    If TotalRegistros(Sql3) = 0 Then
                        Sql = "INSERT INTO albaran_costes(numalbar, numlinea, tipogasto, codcoste, impcoste, importes, unidades, codartic) "
                        Sql = Sql & " VALUES (" & DBLet(Rs!numalbar, "N") & "," & DBSet(Rs1!numlinea, "N") & ",3," & ValorNulo & ","
                        Sql = Sql & DBSet(ComisionesLinea, "N") & "," & DBSet(ComisionesLinea, "N") & ",1," & ValorNulo & ")"
                    Else
                        Sql = "update albaran_costes set impcoste = impcoste + " & DBSet(ComisionesLinea, "N")
                        Sql = Sql & ", importes = importes + " & DBSet(ComisionesLinea, "N")
                        Sql = Sql & " where numalbar = " & DBSet(Rs!numalbar, "N")
                        Sql = Sql & " and numlinea = " & DBSet(Rs1!numlinea, "N")
                        Sql = Sql & " and tipogasto = 3"
                    End If
                    
                    conn.Execute Sql
                
                    Sql = "insert into tmpLineas (numalbar, numlinea, importel) values ("
                    Sql = Sql & DBSet(Rs!numalbar, "N") & "," & DBSet(Rs1!numlinea, "N") & ","
                    Sql = Sql & DBSet(ComisionesLinea, "N") & ")"
                
                    conn.Execute Sql
                
                
                    Rs1.MoveNext
                Wend
                Set Rs1 = Nothing
                
                If RestoComisiones <> 0 Then
                    Sql = "update albaran_costes set impcoste = impcoste + " & DBSet(RestoComisiones, "N")
                    Sql = Sql & ", importes = importes + " & DBSet(RestoComisiones, "N")
                    Sql = Sql & " where numalbar = " & DBLet(Rs!numalbar, "N") & " and numlinea = " & DBSet(Linea, "N")
                    Sql = Sql & " and tipogasto = 3 "
                    
                    conn.Execute Sql
                
                    Sql = "update tmpLineas set importel = importel + " & DBSet(RestoComisiones, "N")
                    Sql = Sql & " where numalbar = " & DBLet(Rs!numalbar, "N") & " and numlinea = " & DBSet(Linea, "N")
                    
                    conn.Execute Sql
                    
                End If
            End If
        Rs.MoveNext
    Wend
    
    InsertarCostesComision = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCostesComision = False
        cadErr = "Error al insertar líneas de Comisiones de albarán: " & Err.Description
    Else
        InsertarCostesComision = True
    End If
End Function


Private Function InsertarCostesComisionLin(cadSQL As String, Optional cadErr As String) As Boolean
'Insertar las lineas de albaran_costes para cada variedad prorrateada con los kilos
Dim Sql As String, Sql2 As String, Sql3 As String
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim NetoTotal As Currency, RestoComisiones As Currency, ComisionesLinea As Currency, PaletsLinea As Currency
Dim Linea As Currency

    On Error GoTo EInsertarAlb

    Sql = "select numalbar, numlinea, portespag from tmpalbaranes where codusu = " & DBSet(vUsu.codigo, "N")
    Sql = Sql & " and portespag <> 0 "
    Sql = Sql & " order by 1,2"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql3 = "select count(*) from albaran_costes where numalbar = " & DBLet(Rs!numalbar, "N")
        Sql3 = Sql3 & " and numlinea = " & DBSet(Rs!numlinea, "N")
        Sql3 = Sql3 & " and tipogasto = 3"
        
        If TotalRegistros(Sql3) = 0 Then
            Sql = "INSERT INTO albaran_costes(numalbar, numlinea, tipogasto, codcoste, impcoste, importes, unidades, codartic) "
            Sql = Sql & " VALUES (" & DBLet(Rs!numalbar, "N") & "," & DBSet(Rs!numlinea, "N") & ",3," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!portespag, "N") & "," & DBSet(Rs!portespag, "N") & ",1," & ValorNulo & ")"
        Else
            Sql = "update albaran_costes set impcoste = impcoste + " & DBSet(Rs!portespag, "N")
            Sql = Sql & ", importes = importes + " & DBSet(Rs!portespag, "N")
            Sql = Sql & " where numalbar = " & DBSet(Rs!numalbar, "N")
            Sql = Sql & " and numlinea = " & DBSet(Rs!numlinea, "N")
            Sql = Sql & " and tipogasto = 3"
        End If
        
        conn.Execute Sql
    
        Sql = "insert into tmpLineas (numalbar, numlinea, importel) values ("
        Sql = Sql & DBSet(Rs!numalbar, "N") & "," & DBSet(Rs!numlinea, "N") & ","
        Sql = Sql & DBSet(Rs!portespag, "N") & ")"
    
        conn.Execute Sql
    
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
                
    InsertarCostesComisionLin = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCostesComisionLin = False
        cadErr = "Error al insertar líneas de Comisiones de lineas de albarán: " & Err.Description
    Else
        InsertarCostesComisionLin = True
    End If
End Function




