VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTransportista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'TABLA: agencias
'   AGENCIAS DE TRANSPORTE
'Variables locales que contienen valores de propiedad
Private mCodTrans As String 'Cod. transportista
Private mNomTrans As String 'Nombre transportista
Private mNomComer As String 'Nombre comercial transportista
Private mDomTrans As String 'Domicilio transportista
Private mCodpobla As String 'cod. postal transportista
Private mPobTrans As String 'poblacion transportista
Private mProTrans As String 'Provincia transportista
Private mNifTrans As String 'NIF transportista
Private mTfn1Trans As String 'Telefono
Private mMai1Trans As String 'EMail de administracion
Private mWwwTrans As String 'Web transportista

Private mCodForpa As String 'cod. forma de pago
Private mDtoPpago As String 'descuento pronto pago
Private mDtoGnral As String 'descuento general
Private mTipoDtos As Byte   'Tipo descuentos: 0=aditivo, 1=Sobre resto
Private mPorReten As String 'porcentaje de retencion

Private mBancoPr As String  'cod. banco propio

'[Monica]22/11/2013: tema iban
Private mIban As String 'iban

Private mCodbanco As String 'cod. banco
Private mCodsucur As String 'cod. sucur
Private mDigContr As String 'digito control
Private mCuentaba As String 'cuenta bancaria

Private mCodmacta As String 'cta contable del transportista

Private mTipoTrans As Byte ' tipo: 0=transportista
                           '       1=comisionista

Private mObservac As String 'Observaciones


'------------------------------------------------
'------------------------------------------------
'Propiedades del modulo CTransportista
'------------------------------------------------
'------------------------------------------------


'**** Codigo del transportista

Public Property Let codigo(ByVal vData As String)
     mCodTrans = vData
End Property

Public Property Get codigo() As String
     codigo = mCodTrans
End Property


'**** Nombre del transportista

Public Property Let Nombre(ByVal vData As String)
     mNomTrans = vData
End Property

Public Property Get Nombre() As String
     Nombre = mNomTrans
End Property

'**** Domicilio del transportista

Public Property Let Domicilio(ByVal vData As String)
     mDomTrans = vData
End Property

Public Property Get Domicilio() As String
     Domicilio = mDomTrans
End Property

'**** CPostal del transportista

Public Property Let CPostal(ByVal vData As String)
     mCodpobla = vData
End Property

Public Property Get CPostal() As String
     CPostal = mCodpobla
End Property

'**** Poblacion del transportista

Public Property Let Poblacion(ByVal vData As String)
     mPobTrans = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mPobTrans
End Property

'**** Provincia del transportista

Public Property Let Provincia(ByVal vData As String)
     mProTrans = vData
End Property

Public Property Get Provincia() As String
     Provincia = mProTrans
End Property

'**** NIF del transportista

Public Property Let NIF(ByVal vData As String)
     mNifTrans = vData
End Property

Public Property Get NIF() As String
     NIF = mNifTrans
End Property

'**** Telefono administracion del transportista

Public Property Let TfnoAdmon(ByVal vData As String)
     mTfn1Trans = vData
End Property

Public Property Get TfnoAdmon() As String
     TfnoAdmon = mTfn1Trans
End Property



'**** EMail administracion del transportista

Public Property Let EMailAdmon(ByVal vData As String)
     mMai1Trans = vData
End Property

Public Property Get EMailAdmon() As String
     EMailAdmon = mMai1Trans
End Property



'**** Web del transportista

Public Property Let WebTrans(ByVal vData As String)
     mWwwTrans = vData
End Property

Public Property Get WebTrans() As String
     WebTrans = mWwwTrans
End Property




'**** Forma de pago del transportista

Public Property Let ForPago(ByVal vData As String)
     mCodForpa = vData
End Property

Public Property Get ForPago() As String
     ForPago = mCodForpa
End Property

'**** Descuento pronto pago del transportista

Public Property Let DtoPPago(ByVal vData As String)
     mDtoPpago = vData
End Property

Public Property Get DtoPPago() As String
     DtoPPago = mDtoPpago
End Property

'**** Descuento general del transportista

Public Property Let DtoGnral(ByVal vData As String)
     mDtoGnral = vData
End Property

Public Property Get DtoGnral() As String
     DtoGnral = mDtoGnral
End Property

'**** Tipo Descuento del transportista

Public Property Let TipoDto(ByVal vData As Byte)
     mTipoDtos = vData
End Property

Public Property Get TipoDto() As Byte
     TipoDto = mTipoDtos
End Property

'**** Porcentaje de retencion del transportista

Public Property Let PorcReten(ByVal vData As String)
     mPorReten = vData
End Property

Public Property Get PorcReten() As String
     PorcReten = mPorReten
End Property




'**** Banco propio del transportista

Public Property Let BancoPropio(ByVal vData As String)
     mBancoPr = vData
End Property

Public Property Get BancoPropio() As String
     BancoPropio = mBancoPr
End Property

'[Monica]22/11/2013: tema iban
'**** Iban del transportista

Public Property Let Iban(ByVal vData As String)
     mIban = vData
End Property

Public Property Get Iban() As String
     Iban = mIban
End Property



'**** Banco del transportista

Public Property Let Banco(ByVal vData As String)
     mCodbanco = vData
End Property

Public Property Get Banco() As String
     Banco = mCodbanco
End Property


'**** Sucursal del transportista

Public Property Let Sucursal(ByVal vData As String)
     mCodsucur = vData
End Property

Public Property Get Sucursal() As String
     Sucursal = mCodsucur
End Property

'**** Digito Control de la cuenta del transportista
Public Property Let DigControl(ByVal vData As String)
     mDigContr = vData
End Property

Public Property Get DigControl() As String
     DigControl = mDigContr
End Property

'**** Cuenta Bancaria del transportista

Public Property Let CuentaBan(ByVal vData As String)
     mCuentaba = vData
End Property

Public Property Get CuentaBan() As String
     CuentaBan = mCuentaba
End Property


'**** Cuenta contable del transportista

Public Property Let CuentaCble(ByVal vData As String)
     mCodmacta = vData
End Property

Public Property Get CuentaCble() As String
     CuentaCble = mCodmacta
End Property

'**** Observaciones del transportista

Public Property Let Observaciones(ByVal vData As String)
     mObservac = vData
End Property

Public Property Get Observaciones() As String
     Observaciones = mObservac
End Property


'**** Tipo Descuento del transportista

Public Property Let EsComisionista(ByVal vData As Byte)
     mTipoTrans = vData
End Property

Public Property Get EsComisionista() As Byte
     EsComisionista = mTipoTrans
End Property




'------------------------------------------------
'------------------------------------------------
'Procedimientos del modulo CTransportista
'------------------------------------------------
'------------------------------------------------

Public Function Existe(vCodTrans As String) As Boolean
'Comprueba si existe el transportista en la BD
Dim devuelve As String

    On Error GoTo EExiste
    
    devuelve = DevuelveDesdeBDNew(cAgro, "agencias", "codtrans", "codtrans", vCodTrans, "N")
    If devuelve = "" Then
        Existe = False
        devuelve = "No existe el Transportista " & vCodTrans & ". "
        MsgBox devuelve, vbExclamation
    Else
        Existe = True
    End If
    
EExiste:
    If Err.Number <> 0 Then Existe = False
End Function


Public Function LeerDatos(vCodTrans As String) As Boolean
'Leer los datos de un tranportista dado
'Lee de la BD: Ariges, Tabla: agencias
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim Mens As String
Dim SQL As String

    On Error GoTo ELeer
    LeerDatos = False
    
    SQL = "SELECT codtrans,nomtrans,codpobla,domtrans,pobtrans,protrans,ciftrans,teltrans1,maitrans1,wwwtrans,codforpa,codbanpr,codbanco,codsucur,digcontr,cuentaba,codmacta,obstrans, porcereten, tipo, iban"
    SQL = SQL & " FROM agencias "
    SQL = SQL & " WHERE codtrans=" & vCodTrans
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    If RS.EOF Then
        LeerDatos = False
    Else
        Mens = "Datos Transportista"
        mCodTrans = CStr(RS!codTrans)
        
        mNomTrans = RS!NomTrans
        mDomTrans = DBLet(RS!domtrans, "T")
        mCodpobla = DBLet(RS!codPobla, "T")
        mPobTrans = DBLet(RS!pobtrans, "T")
        mProTrans = DBLet(RS!protrans, "T")
        mNifTrans = RS!ciftrans
        mTfn1Trans = DBLet(RS!teltrans1, "T")
        mMai1Trans = DBLet(RS!maitrans1, "T")
        mWwwTrans = DBLet(RS!wwwtrans, "T")
        
        mCodForpa = CStr(RS!Codforpa)
        mDtoPpago = 0
        mDtoGnral = 0
        mTipoDtos = 0
        mPorReten = CStr(DBLet(RS!porcereten, "N"))
        
        mBancoPr = CStr(RS!codBanpr)
        '[Monica]22/11/2013: tema iban
        mIban = DBLet(RS!Iban, "T")
        mCodbanco = CStr(DBLet(RS!codbanco, "N"))
        '[Monica]21/11/2014: le damos formato
        If mCodbanco <> "" And ComprobarCero(mCodbanco) <> 0 Then mCodbanco = Format(CInt(mCodbanco), "0000")
        
        mCodsucur = CStr(DBLet(RS!codsucur, "N"))
        '[Monica]21/11/2014: le damos formato
        If mCodsucur <> "" And ComprobarCero(mCodsucur) <> 0 Then mCodsucur = Format(CInt(mCodsucur), "0000")
        
        mDigContr = DBLet(RS!digcontr, "T")
        mCuentaba = DBLet(RS!cuentaba, "T")
        
        mCodmacta = DBLet(RS!Codmacta, "T")
        mTipoTrans = CStr(DBLet(RS!Tipo, "N"))
        
        mObservac = DBLet(RS!obstrans, "T")
        
        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set RS = Nothing
    LeerDatos = False
End Function

