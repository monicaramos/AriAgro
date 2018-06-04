VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParamAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

'Variables locales que contienen valores de propiedad
'Variables asociadas a cada campo de la tabla de parametros SPARAM

'Datos de la Contabilidad
Private mvarSerConta As String 'Servidor BD de Contabilidad
Private mvarUsuConta As String 'Usuario BD de Contabilidad
Private mvarPasConta As String 'Password BD de Contabilidad
Private mvarNumConta As Byte 'Numero BD de Contabilidad
Private mvarLimPesoBrutoCMR As Long 'Limite Peso Bruto de CMR
Private mvarCodIvaTrans As Byte 'Codigo Iva Factura de transporte
Private mvarCodIvaNormal As Byte 'Codigo Iva de venta de facturas a cuenta
Private mvarCodIvaExento As Byte 'Codigo Iva Cliente Exento
Private mvarCodIvaRecargo As Byte 'Codigo Iva Cliente Recargo

Private mvarWebSoporte As String  'Web Soporte
Private mvarDireMail As String 'direccion mail
Private mvarSmtphost As String 'Host
Private mvarSmtpUser As String 'User
Private mvarSmtpPass As String 'Password

Private mvarTipoPrecio As Byte 'Tipo precio
Private mvarInventar As Boolean  'Si se realiza Inventario ordenando listados por Proveedor o no
Private mvarCtrStock As Boolean   'Si hay Control de Stock (S/N)
Private mvarPaseAlbarAgrupCalib As Boolean  'Si se detalla una linea de variedad por
                                            'cada calibre en el pase de pedidos a albaran partiendo
                                            ' de palets.

Private mvarPaseRefLineaAlb As Boolean  ' si esta marcado pasamos la refencia de linea de albaran sino pasamos la
                                        ' referencia de cliente de cabecera de albaran al arimoney
                                        
Private mvarTipoCalculoComision As Byte  ' Tipo Calculo de dtos y comisiones en facturas de cliente
                                            ' 0 = dtos y comisiones sobre importe bruto
                                            ' 1 = dtos sobre importe bruto, comisiones sobre kilos reales

'[Monica]25/04/2012: nuevo
Private mvarNumeroGGN As Double ' Nro de GGN

' facturacion
Private mvarDiaPago1 As Byte 'Dia de pago 1 para compras
Private mvarDiaPago2 As Byte 'Dia de pago 2 para compras
Private mvarDiaPago3 As Byte 'Dia de pago 3 para compras
Private mvarMesnogir As Byte 'Mes a no girar para compras

Private mvarCtaTraReten As String 'Cuenta Contable Transporte de Retencion
Private mvarCtaComReten As String 'Cuenta Contable de Retencion Facturas de Comisionistas
Private mvarCtaAboTrans As String 'Cuenta Contable Abono de Transporte
Private mvarCtaVentasFraACta As String 'Cuenta Contable de ventas para facturas a cuenta

Private mvarCCosteFraACta As String 'Centro de coste para facturas a cuenta


Private mvarNroLote As Integer 'Nro Lote - grabacion fichero cheps
Private mvarNroCheps As Currency 'Nro cliente cheps - grabacion fichero cheps
Private mvarNroFiche As Currency 'Nro fichero cheps - grabacion fichero cheps
Private mvarAlmacen As Integer 'codigo de almacen por defecto para albaranes

Private mvarNroPolizaExp As String 'Nro de Poliza de Expediente

' aridoc
Private mvarHayAridoc As Byte  'Hay Aridoc
Private mvarCarpetaAlb As Integer 'Carpeta Albaranes
Private mvarCarpetaFac As Integer 'Carpeta Facturas
Private mvarCarpetaRecAlmacen As Integer 'Carpeta Recibos de Almacen
Private mvarCarpetaRecCampo As Integer 'Carpeta Recibos de Campo
Private mvarExtension As Byte  'Extension del documento
Private mvarC1Albaran As Byte  'Campo 1 Albaran
Private mvarC2Albaran As Byte  'Campo 2 Albaran
Private mvarC3Albaran As Byte  'Campo 3 Albaran
Private mvarC4Albaran As Byte  'Campo 4 Albaran
Private mvarC1Factura As Byte  'Campo 1 Factura
Private mvarC2Factura As Byte  'Campo 2 Factura
Private mvarC3Factura As Byte  'Campo 3 Factura
Private mvarC4Factura As Byte  'Campo 4 Factura

Private mvarC1Recibo As Byte  'Campo 1 Recibo
Private mvarC2Recibo As Byte  'Campo 2 Recibo
Private mvarC3Recibo As Byte  'Campo 3 Recibo
Private mvarC4Recibo As Byte  'Campo 4 Recibo


'edi
Private mvarCodigoEdi As String 'Codigo edi del vendedor
Private mvarRegMercantil As String 'Registro Mercantil emisor factura
Private mvarPathEdicom As String 'Path de ficheros generados para edicom

'varios
Private mvarPorcDesv As Currency 'Porcentaje de desviacion para el calculo de gastos
Private mvarPortesKiloCaja As Byte  'Portes por kilo o caja


Private mvarCooperativa As Byte ' 0 = Catadau
                                ' 1 = Valsur
                                ' 2 = Picassent
                                ' 3 = Mogente
                                ' 4 = Alzira
                                ' 5 = Castelduc
                                ' 6 = Macastre
                                ' 7 = Quatretonda
                                ' 8 = Utxera
                                ' 9 = Natural
                                ' 10= Escalona
                                ' 11= Belgida
                                ' 12= Montifrut
                                ' 14= Bolbaite
                                ' 15= IMG

Private mvarCodTipomAlb As String 'Tipo de moviemiento de albaranes de salida

' solapa de Horto
Private mvarText1CMR13 As String ' Texto 1 para el CMR seccion 13
Private mvarText2CMR13 As String ' Texto 2 para el CMR seccion 13

Private mvarClienteVtas As String ' cliente de albaranes de venta

Private mvarHayCCostes As Byte  'Hay Control de costes

Private mvarEnvioDesdeOutlook As Boolean  'Si envia por outlook
Private mvarExeEnvioMail As String


' Costes dentro de la solapa Otros
Private mvarPathFichadas As String 'Path de ficheros de fichadas de costes

Private mvarPathFacturaE As String  'FacturaE

Private mvarClaveAcceso As String ' Clave de acceso

Private mvarHayAnecoop As Byte  'Hay Integracion Anecoop

Private mvarCodAnecoop As String ' Cliente de Anecoop
Private mvarEnvAnecoop As String ' Envase de Anecoop

Private mvarContabilidadNueva As Boolean

Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes





'------------------------------------------------
'Propiedades del modulo CParamAplic
'------------------------------------------------

'servidor conta
Public Property Let ServidorConta(ByVal vData As String)
     mvarSerConta = vData
End Property

Public Property Get ServidorConta() As String
     ServidorConta = mvarSerConta
End Property

Public Property Let UsuarioConta(ByVal vData As String)
     mvarUsuConta = vData
End Property

Public Property Get UsuarioConta() As String
     UsuarioConta = mvarUsuConta
End Property

Public Property Let PasswordConta(ByVal vData As String)
     mvarPasConta = vData
End Property

Public Property Get PasswordConta() As String
     PasswordConta = mvarPasConta
End Property

Public Property Let NumeroConta(ByVal vData As Byte)
     mvarNumConta = vData
End Property

Public Property Get NumeroConta() As Byte
     NumeroConta = mvarNumConta
End Property

Public Property Let LimPesoCMR(ByVal vData As Long)
     mvarLimPesoBrutoCMR = vData
End Property

Public Property Get LimPesoCMR() As Long
     LimPesoCMR = mvarLimPesoBrutoCMR
End Property

Public Property Let WebSoporte(ByVal vData As String)
     mvarWebSoporte = vData
End Property

Public Property Get WebSoporte() As String
     WebSoporte = mvarWebSoporte
End Property

Public Property Let DireMail(ByVal vData As String)
     mvarDireMail = vData
End Property

Public Property Get DireMail() As String
     DireMail = mvarDireMail
End Property

Public Property Let Smtphost(ByVal vData As String)
     mvarSmtphost = vData
End Property

Public Property Get Smtphost() As String
     Smtphost = mvarSmtphost
End Property

Public Property Let SmtpUser(ByVal vData As String)
     mvarSmtpUser = vData
End Property

Public Property Get SmtpUser() As String
     SmtpUser = mvarSmtpUser
End Property

Public Property Let Smtppass(ByVal vData As String)
     mvarSmtpPass = vData
End Property

Public Property Get Smtppass() As String
     Smtppass = mvarSmtpPass
End Property

Public Property Let TipoPrecio(ByVal vData As Byte)
     mvarTipoPrecio = vData
End Property

Public Property Get TipoPrecio() As Byte
     TipoPrecio = mvarTipoPrecio
End Property

Public Property Let InventarioxProv(ByVal vData As Boolean)
     mvarInventar = vData
End Property

Public Property Get InventarioxProv() As Boolean
     InventarioxProv = mvarInventar
End Property

Public Property Let ControlStock(ByVal vData As Boolean)
     mvarCtrStock = vData
End Property

Public Property Get ControlStock() As Boolean
     ControlStock = mvarCtrStock
End Property

Public Property Let TipoCalculoComision(ByVal vData As Byte)
     mvarTipoCalculoComision = vData
End Property

Public Property Get TipoCalculoComision() As Byte
     TipoCalculoComision = mvarTipoCalculoComision
End Property


Public Property Let PaseAlbarAgrupCalib(ByVal vData As Boolean)
     mvarPaseAlbarAgrupCalib = vData
End Property

Public Property Get PaseAlbarAgrupCalib() As Boolean
     PaseAlbarAgrupCalib = mvarPaseAlbarAgrupCalib
End Property


Public Property Let PaseRefLineaAlb(ByVal vData As Boolean)
     mvarPaseRefLineaAlb = vData
End Property

Public Property Get PaseRefLineaAlb() As Boolean
     PaseRefLineaAlb = mvarPaseRefLineaAlb
End Property


Public Property Let DiaPago1(ByVal vData As Byte)
     mvarDiaPago1 = vData
End Property

Public Property Get DiaPago1() As Byte
     DiaPago1 = mvarDiaPago1
End Property

Public Property Let DiaPago2(ByVal vData As Byte)
     mvarDiaPago2 = vData
End Property

Public Property Get DiaPago2() As Byte
     DiaPago2 = mvarDiaPago2
End Property

Public Property Let DiaPago3(ByVal vData As Byte)
     mvarDiaPago3 = vData
End Property

Public Property Get DiaPago3() As Byte
     DiaPago3 = mvarDiaPago3
End Property


Public Property Let MesNoGirar(ByVal vData As Byte)
     mvarMesnogir = vData
End Property

Public Property Get MesNoGirar() As Byte
     MesNoGirar = mvarMesnogir
End Property

Public Property Let CodIvaTrans(ByVal vData As Byte)
     mvarCodIvaTrans = vData
End Property

Public Property Get CodIvaTrans() As Byte
     CodIvaTrans = mvarCodIvaTrans
End Property

Public Property Let CtaTraReten(ByVal vData As String)
     mvarCtaTraReten = vData
End Property

Public Property Get CtaTraReten() As String
     CtaTraReten = mvarCtaTraReten
End Property

Public Property Let CtaAboTrans(ByVal vData As String)
     mvarCtaAboTrans = vData
End Property

Public Property Get CtaAboTrans() As String
     CtaAboTrans = mvarCtaAboTrans
End Property

Public Property Let CtaComReten(ByVal vData As String)
     mvarCtaComReten = vData
End Property

Public Property Get CtaComReten() As String
     CtaComReten = mvarCtaComReten
End Property

Public Property Let CtaVentasFraACta(ByVal vData As String)
     mvarCtaVentasFraACta = vData
End Property

Public Property Get CtaVentasFraACta() As String
     CtaVentasFraACta = mvarCtaVentasFraACta
End Property

Public Property Let CCosteFraACta(ByVal vData As String)
     mvarCCosteFraACta = vData
End Property

Public Property Get CCosteFraACta() As String
     CCosteFraACta = mvarCCosteFraACta
End Property





' Factura de ventas
Public Property Let CodIvaExento(ByVal vData As Byte)
     mvarCodIvaExento = vData
End Property

Public Property Get CodIvaExento() As Byte
     CodIvaExento = mvarCodIvaExento
End Property

Public Property Let CodIvaNormal(ByVal vData As Byte)
     mvarCodIvaNormal = vData
End Property

Public Property Get CodIvaNormal() As Byte
     CodIvaNormal = mvarCodIvaNormal
End Property

Public Property Let CodIvaRecargo(ByVal vData As Byte)
     mvarCodIvaRecargo = vData
End Property

Public Property Get CodIvaRecargo() As Byte
     CodIvaRecargo = mvarCodIvaRecargo
End Property

' Aridoc
Public Property Let HayAridoc(ByVal vData As Byte)
     mvarHayAridoc = vData
End Property

Public Property Get HayAridoc() As Byte
     HayAridoc = mvarHayAridoc
End Property

Public Property Let CarpetaAlb(ByVal vData As Integer)
     mvarCarpetaAlb = vData
End Property

Public Property Get CarpetaAlb() As Integer
     CarpetaAlb = mvarCarpetaAlb
End Property

Public Property Let CarpetaFac(ByVal vData As Integer)
     mvarCarpetaFac = vData
End Property

Public Property Get CarpetaFac() As Integer
     CarpetaFac = mvarCarpetaFac
End Property

Public Property Let CarpetaRecAlmacen(ByVal vData As Integer)
     mvarCarpetaRecAlmacen = vData
End Property

Public Property Get CarpetaRecAlmacen() As Integer
     CarpetaRecAlmacen = mvarCarpetaRecAlmacen
End Property

Public Property Let CarpetaRecCampo(ByVal vData As Integer)
     mvarCarpetaRecCampo = vData
End Property

Public Property Get CarpetaRecCampo() As Integer
     CarpetaRecCampo = mvarCarpetaRecCampo
End Property




Public Property Let Extension(ByVal vData As Byte)
     mvarExtension = vData
End Property

Public Property Get Extension() As Byte
     Extension = mvarExtension
End Property


Public Property Let C1Albaran(ByVal vData As Byte)
     mvarC1Albaran = vData
End Property

Public Property Get C1Albaran() As Byte
     C1Albaran = mvarC1Albaran
End Property

Public Property Let C2Albaran(ByVal vData As Byte)
     mvarC2Albaran = vData
End Property

Public Property Get C2Albaran() As Byte
     C2Albaran = mvarC2Albaran
End Property

Public Property Let C3Albaran(ByVal vData As Byte)
     mvarC3Albaran = vData
End Property

Public Property Get C3Albaran() As Byte
     C3Albaran = mvarC3Albaran
End Property

Public Property Let C4Albaran(ByVal vData As Byte)
     mvarC4Albaran = vData
End Property

Public Property Get C4Albaran() As Byte
     C4Albaran = mvarC4Albaran
End Property


Public Property Let C1Factura(ByVal vData As Byte)
     mvarC1Factura = vData
End Property

Public Property Get C1Factura() As Byte
     C1Factura = mvarC1Factura
End Property

Public Property Let C2Factura(ByVal vData As Byte)
     mvarC2Factura = vData
End Property

Public Property Get C2Factura() As Byte
     C2Factura = mvarC2Factura
End Property

Public Property Let C3Factura(ByVal vData As Byte)
     mvarC3Factura = vData
End Property

Public Property Get C3Factura() As Byte
     C3Factura = mvarC3Factura
End Property

Public Property Let C4Factura(ByVal vData As Byte)
     mvarC4Factura = vData
End Property

Public Property Get C4Factura() As Byte
     C4Factura = mvarC4Factura
End Property


Public Property Let C1Recibo(ByVal vData As Byte)
     mvarC1Recibo = vData
End Property

Public Property Get C1Recibo() As Byte
     C1Recibo = mvarC1Recibo
End Property

Public Property Let C2Recibo(ByVal vData As Byte)
     mvarC2Recibo = vData
End Property

Public Property Get C2Recibo() As Byte
     C2Recibo = mvarC2Recibo
End Property

Public Property Let C3Recibo(ByVal vData As Byte)
     mvarC3Recibo = vData
End Property

Public Property Get C3Recibo() As Byte
     C3Recibo = mvarC3Recibo
End Property

Public Property Let C4Recibo(ByVal vData As Byte)
     mvarC4Recibo = vData
End Property

Public Property Get C4Recibo() As Byte
     C4Recibo = mvarC4Recibo
End Property




Public Property Let NroLote(ByVal vData As Integer)
     mvarNroLote = vData
End Property

Public Property Get NroLote() As Integer
     NroLote = mvarNroLote
End Property

Public Property Let NroCheps(ByVal vData As Currency)
     mvarNroCheps = vData
End Property

Public Property Get NroCheps() As Currency
     NroCheps = mvarNroCheps
End Property

Public Property Let NroFiche(ByVal vData As Currency)
     mvarNroFiche = vData
End Property

Public Property Get NroFiche() As Currency
     NroFiche = mvarNroFiche
End Property

'Codigo edi
Public Property Let CodigoEdi(ByVal vData As String)
     mvarCodigoEdi = vData
End Property

Public Property Get CodigoEdi() As String
     CodigoEdi = mvarCodigoEdi
End Property

'Registro Mercantil
Public Property Let RegMercantil(ByVal vData As String)
     mvarRegMercantil = vData
End Property

Public Property Get RegMercantil() As String
     RegMercantil = mvarRegMercantil
End Property

'Path Edicom
Public Property Let PathEdicom(ByVal vData As String)
     mvarPathEdicom = vData
End Property

Public Property Get PathEdicom() As String
     PathEdicom = mvarPathEdicom
End Property

Public Property Let PorcDesvCostes(ByVal vData As Currency)
     mvarPorcDesv = vData
End Property

Public Property Get PorcDesvCostes() As Currency
     PorcDesvCostes = mvarPorcDesv
End Property


Public Property Let PortesKiloCaja(ByVal vData As Byte)
     mvarPortesKiloCaja = vData
End Property

Public Property Get PortesKiloCaja() As Byte
     PortesKiloCaja = mvarPortesKiloCaja
End Property

Public Property Let Almacen(ByVal vData As Integer)
     mvarAlmacen = vData
End Property

Public Property Get Almacen() As Integer
     Almacen = mvarAlmacen
End Property

Public Property Let NroPolizaExp(ByVal vData As String)
     mvarNroPolizaExp = vData
End Property

Public Property Get NroPolizaExp() As String
     NroPolizaExp = mvarNroPolizaExp
End Property

Public Property Let Cooperativa(ByVal vData As Byte)
     mvarCooperativa = vData
End Property

Public Property Get Cooperativa() As Byte
     Cooperativa = mvarCooperativa
End Property


Public Property Let CodTipomAlb(ByVal vData As String)
     mvarCodTipomAlb = vData
End Property

Public Property Get CodTipomAlb() As String
     CodTipomAlb = mvarCodTipomAlb
End Property

' Texto 1 para CMR seccion 13
Public Property Get Text1CMR13() As String
     Text1CMR13 = mvarText1CMR13
End Property

Public Property Let Text1CMR13(ByVal vData As String)
     mvarText1CMR13 = vData
End Property

' Texto 2 para CMR seccion 13
Public Property Get Text2CMR13() As String
     Text2CMR13 = mvarText2CMR13
End Property

Public Property Let Text2CMR13(ByVal vData As String)
     mvarText2CMR13 = vData
End Property


Public Property Let ClienteVtas(ByVal vData As String)
     mvarClienteVtas = vData
End Property

Public Property Get ClienteVtas() As String
     ClienteVtas = mvarClienteVtas
End Property


' Hay control de costes
Public Property Let HayCCostes(ByVal vData As Byte)
     mvarHayCCostes = vData
End Property

Public Property Get HayCCostes() As Byte
     HayCCostes = mvarHayCCostes
End Property

' Envio de email por outlook o arigesmail

Public Property Get EnvioDesdeOutlook() As Boolean
     EnvioDesdeOutlook = mvarEnvioDesdeOutlook
End Property

Public Property Let EnvioDesdeOutlook(ByVal vData As Boolean)
     mvarEnvioDesdeOutlook = vData
End Property

Public Property Get ExeEnvioMail() As String
     ExeEnvioMail = mvarExeEnvioMail
End Property

Public Property Let ExeEnvioMail(ByVal vData As String)
     mvarExeEnvioMail = vData
End Property

Public Property Let NumeroGGN(ByVal vData As Double)
     mvarNumeroGGN = vData
End Property

Public Property Get NumeroGGN() As Double
     NumeroGGN = mvarNumeroGGN
End Property


'Path Fichadas
Public Property Let PathFichadas(ByVal vData As String)
     mvarPathFichadas = vData
End Property

Public Property Get PathFichadas() As String
     PathFichadas = mvarPathFichadas
End Property


'Facturacion Electrónica FacturaE
Public Property Get PathFacturaE() As String
     PathFacturaE = mvarPathFacturaE
End Property

Public Property Let PathFacturaE(ByVal vData As String)
     mvarPathFacturaE = vData
End Property


Public Property Let ClaveAcceso(ByVal vData As String)
     mvarClaveAcceso = vData
End Property

Public Property Get ClaveAcceso() As String
     ClaveAcceso = mvarClaveAcceso
End Property

' Aridoc
Public Property Let HayAnecoop(ByVal vData As Byte)
     mvarHayAnecoop = vData
End Property

Public Property Get HayAnecoop() As Byte
     HayAnecoop = mvarHayAnecoop
End Property

Public Property Let CodAnecoop(ByVal vData As String)
     mvarCodAnecoop = vData
End Property

Public Property Get CodAnecoop() As String
     CodAnecoop = mvarCodAnecoop
End Property

Public Property Let EnvAnecoop(ByVal vData As String)
     mvarEnvAnecoop = vData
End Property

Public Property Get EnvAnecoop() As String
     EnvAnecoop = mvarEnvAnecoop
End Property

'[Monica]03/11/2016: si hay contabilidad nueva
Public Property Get ContabilidadNueva() As Boolean
     ContabilidadNueva = mvarContabilidadNueva
End Property


'******
'------------------------------------------------
'FUNCIONES del modulo Cparametros
'------------------------------------------------

'****   LEER  ***
Public Function Leer() As Byte
Dim Rs As ADODB.Recordset

On Error GoTo Err3

    SQL = "SELECT * FROM sparam "
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic

    If Rs.EOF Then
        Leer = 1
    Else
        
        'Datos Contabilidad
        mvarSerConta = DBLet(Rs!serConta, "T") 'servidor conta
        mvarUsuConta = DBLet(Rs!UsuConta, "T") 'usuario conta
        mvarPasConta = DBLet(Rs!pasconta, "T") 'password
        mvarNumConta = DBLet(Rs!NumConta, "N") 'numero conta
        mvarWebSoporte = DBLet(Rs!WebSoporte, "T") 'web soporte
        mvarDireMail = DBLet(Rs!DireMail, "T") 'mail
        mvarSmtphost = DBLet(Rs!Smtphost, "T") 'host
        mvarSmtpUser = DBLet(Rs!SmtpUser, "T") 'user
        mvarSmtpPass = DBLet(Rs!Smtppass, "T") 'pass
        mvarTipoPrecio = DBLet(Rs!TipoPrecio, "N") ' tipo de precio
        mvarInventar = DBLet(Rs!inventar, "B") 'inventario por proveedor o no
        mvarCtrStock = Rs!CtrStock  'Si hay Control de Stock (S/N)
        mvarPaseAlbarAgrupCalib = Rs!PaseAlbarAgrupCalib  'Si se detalla una linea de variedad por calibre
        mvarPaseRefLineaAlb = Rs!PaseRefLineaAlb  'Si se pasa la referencia de linea de albaran al arimoney
        mvarTipoCalculoComision = DBLet(Rs!calculocomision, "N") 'Tipo de prorrateo de comision en factura cliente
        
        mvarDiaPago1 = DBLet(Rs!DiaPago1, "N") 'dia pago 1 para compras
        mvarDiaPago2 = DBLet(Rs!DiaPago2, "N") 'dia pago 2 para compras
        mvarDiaPago3 = DBLet(Rs!DiaPago3, "N") 'dia pago 3 para compras
        mvarMesnogir = DBLet(Rs!mesnogir, "N") 'mes a no girar para compras
        
        mvarLimPesoBrutoCMR = DBLet(Rs!limpesobrutcmr, "N") 'limite de peso bruto para el listado CMR
        mvarCodIvaTrans = DBLet(Rs!CodIvaTrans, "N") 'codigo iva de factura de transporte
        
        mvarCtaTraReten = DBLet(Rs!CtaTraReten, "T") 'cuenta contable de transporte de retencion
        mvarCtaComReten = DBLet(Rs!CtaComReten, "T") 'cuenta contable de retencion de facturas de comisionistas
        mvarCtaAboTrans = DBLet(Rs!CtaAboTrans, "T") 'cuenta contable de abono de transporte
        mvarCtaVentasFraACta = DBLet(Rs!CtaVentasFraACta, "T") 'cuenta contable de ventas para factura a cuenta

        mvarCCosteFraACta = DBLet(Rs!CCosteFraACta, "T") 'centro de coste para factura a cuenta


        mvarCodIvaExento = DBLet(Rs!CodIvaExento, "N") 'codigo iva de factura de exento
        mvarCodIvaNormal = DBLet(Rs!CodIvaNormal, "N") 'codigo iva de factura de venta a cuenta
        mvarCodIvaRecargo = DBLet(Rs!CodIvaRecargo, "N") 'codigo iva de factura con recargo de equivalencia
        
        'aridoc
        mvarHayAridoc = DBLet(Rs!HayAridoc, "N") ' hay aridoc
        mvarCarpetaAlb = DBLet(Rs!codcarpetaalb, "N") 'carpeta albaran
        mvarCarpetaFac = DBLet(Rs!codcarpetafac, "N") 'carpeta factura
        mvarCarpetaRecAlmacen = DBLet(Rs!codcarpetarecalm, "N") 'carpeta recibo almacen
        mvarCarpetaRecCampo = DBLet(Rs!codcarpetareccamp, "N") 'carpeta recibo campo
        
        mvarExtension = DBLet(Rs!codextension, "N") 'extension documento
        mvarC1Albaran = DBLet(Rs!c1albaridoc, "N") 'campo 1 albaran
        mvarC2Albaran = DBLet(Rs!c2albaridoc, "N") 'campo 2 albaran
        mvarC3Albaran = DBLet(Rs!c3albaridoc, "N") 'campo 3 albaran
        mvarC4Albaran = DBLet(Rs!c4albaridoc, "N") 'campo 4 albaran
        mvarC1Factura = DBLet(Rs!c1facaridoc, "N") 'campo 1 factura
        mvarC2Factura = DBLet(Rs!c2facaridoc, "N") 'campo 2 factura
        mvarC3Factura = DBLet(Rs!c3facaridoc, "N") 'campo 3 factura
        mvarC4Factura = DBLet(Rs!c4facaridoc, "N") 'campo 4 factura
        mvarC1Recibo = DBLet(Rs!c1recaridoc, "N") 'campo 1 recibo
        mvarC2Recibo = DBLet(Rs!c2recaridoc, "N") 'campo 2 recibo
        mvarC3Recibo = DBLet(Rs!c3recaridoc, "N") 'campo 3 recibo
        mvarC4Recibo = DBLet(Rs!c4recaridoc, "N") 'campo 4 recibo
        
        mvarNroLote = DBLet(Rs!NroLote, "N") 'nro lote para grabacion fichero chep
        mvarNroCheps = DBLet(Rs!NroCheps, "N")  'nro cliente chep para grabacion fichero chep
        mvarNroFiche = DBLet(Rs!NroFiche, "N")  'nro fichero chep para grabacion fichero chep
        mvarAlmacen = DBLet(Rs!codAlmac, "N")  'codigo de almacen por defecto
        
        mvarCodigoEdi = DBLet(Rs!CodigoEdi, "T") 'codigo edi
        mvarRegMercantil = DBLet(Rs!RegMercantil, "T") 'registro mercantil
        mvarPathEdicom = DBLet(Rs!PathEdicom, "T") 'path edicom
        
        mvarPorcDesv = DBLet(Rs!porcdesvia, "N") 'porcentaje de desviacion de calculo de costes reales
        mvarPortesKiloCaja = DBLet(Rs!PortesKiloCaja, "N") 'calculo de portes por kilo o caja
        
        mvarNroPolizaExp = DBLet(Rs!NroPolizaExp, "T") 'nro de poliza del expediente
        
        mvarCodTipomAlb = DBLet(Rs!CodTipomAlb, "T") 'tipo de movimiento de albaranes de salida
        
        '[Monica] 23/04/2010 : necesitamos la cooperativa
        mvarCooperativa = DevuelveValor("select cooperativa from rparam")
                
        mvarText1CMR13 = DBLet(Rs!Text1CMR13, "T")  ' texto de impresion de cmr para seccion 13
        mvarText2CMR13 = DBLet(Rs!Text2CMR13, "T")  ' texto de impresion de cmr para seccion 13
                
        mvarClienteVtas = DBLet(Rs!CodClien, "N")  ' cliente de ventas
        
        mvarHayCCostes = DBLet(Rs!HayCCostes, "N") ' hay control de costes
        
        mvarEnvioDesdeOutlook = DBLet(Rs!EnvioDesdeOutlook, "N") = 1
        mvarExeEnvioMail = DBLet(Rs!arigesmail, "T")
        
        mvarNumeroGGN = DBLet(Rs!NumeroGGN, "N") ' Numero GGN
        
        mvarPathFichadas = DBLet(Rs!PathFichadas, "T") 'path fichadas
        
        'FacturaE
        mvarPathFacturaE = DBLet(Rs!PathFacturaE, "T")
        
        mvarClaveAcceso = DevuelveValor("select claveacceso from rparam")
        
        mvarHayAnecoop = DBLet(Rs!HayAnecoop, "N") ' hay integracion anecoop
        mvarCodAnecoop = DBLet(Rs!CodAnecoop, "N") ' código de socio anecoop
        mvarEnvAnecoop = DBLet(Rs!EnvAnecoop, "T") ' envase de anecoop

        '[Monica]03/11/2016: nueva contabilidad
        mvarContabilidadNueva = DevuelveValor("select nuevacontabilidad from rparam") = 1

        Leer = 0
        
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Rs.Close
    Set Rs = Nothing
    Leer = 1
End Function


Public Function Modificar() As Byte
'Modifica los parametros de la Aplicación
'Modifica la Tabla: sparam, BD: Arigasol
On Error GoTo EModificar

    SQL = "UPDATE sparam SET "
    SQL = SQL & " serconta= '" & mvarSerConta & "'"
    SQL = SQL & ", usuconta= " & DBSet(mvarUsuConta, "T")
    SQL = SQL & ", pasconta= " & DBSet(mvarPasConta, "T")
    SQL = SQL & ", numconta= " & DBSet(mvarNumConta, "N", "S")
    SQL = SQL & ", websoporte= " & DBSet(mvarWebSoporte, "T")
    SQL = SQL & ", diremail= " & DBSet(mvarDireMail, "T")
    SQL = SQL & ", smtphost= " & DBSet(mvarSmtphost, "T")
    SQL = SQL & ", smtpuser= " & DBSet(mvarSmtpUser, "T")
    SQL = SQL & ", smtppass= " & DBSet(mvarSmtpPass, "T")
    SQL = SQL & ", tipoprecio= " & DBSet(mvarTipoPrecio, "N")
    SQL = SQL & ", inventar= " & DBSet(mvarInventar, "B")
    SQL = SQL & ", ctrstock = " & DBSet(mvarCtrStock, "B")
    SQL = SQL & ", pasealbaragrupcalib = " & DBSet(mvarPaseAlbarAgrupCalib, "B")
    SQL = SQL & ", pasereflineaalb = " & DBSet(mvarPaseRefLineaAlb, "B")
    SQL = SQL & ", diapago1= " & DBSet(mvarDiaPago1, "N")
    SQL = SQL & ", diapago2= " & DBSet(mvarDiaPago2, "N")
    SQL = SQL & ", diapago3= " & DBSet(mvarDiaPago3, "N")
    SQL = SQL & ", mesnogir= " & DBSet(mvarMesnogir, "N")
    SQL = SQL & ", limpesobrutcmr= " & DBSet(mvarLimPesoBrutoCMR, "N")
    SQL = SQL & ", codivatrans= " & DBSet(mvarCodIvaTrans, "N")
    SQL = SQL & ", ctatrareten= " & DBSet(mvarCtaTraReten, "T")
    SQL = SQL & ", ctacomreten= " & DBSet(mvarCtaComReten, "T")
    SQL = SQL & ", ctaabotrans= " & DBSet(mvarCtaAboTrans, "T")
    SQL = SQL & ", ctaventasfraacta= " & DBSet(mvarCtaVentasFraACta, "T")
    SQL = SQL & ", ccostefraacta= " & DBSet(mvarCCosteFraACta, "T") ' centro de coste de facturas a cuenta
    SQL = SQL & ", codivanormal= " & DBSet(mvarCodIvaNormal, "N")
    SQL = SQL & ", codivaexento= " & DBSet(mvarCodIvaExento, "N")
    SQL = SQL & ", codivarecargo= " & DBSet(mvarCodIvaRecargo, "N")
    SQL = SQL & ", nrolote= " & DBSet(mvarNroLote, "N")
    SQL = SQL & ", nrocheps= " & DBSet(mvarNroCheps, "N")
    SQL = SQL & ", nrofiche= " & DBSet(mvarNroFiche, "N")
    SQL = SQL & ", codalmac= " & DBSet(mvarAlmacen, "N")
    SQL = SQL & ", nropolizaexp= " & DBSet(mvarNroPolizaExp, "T")
    SQL = SQL & ", numeroggn= " & DBSet(mvarNumeroGGN, "N")
    
    'aridoc
    SQL = SQL & ", hayaridoc= " & DBSet(mvarHayAridoc, "N") ' hay aridoc
    SQL = SQL & ", codcarpetaalb = " & DBSet(mvarCarpetaAlb, "N") 'carpeta albaran
    SQL = SQL & ", codcarpetafac = " & DBSet(mvarCarpetaFac, "N") 'carpeta factura
    SQL = SQL & ", codcarpetarecalm = " & DBSet(mvarCarpetaRecAlmacen, "N") 'carpeta recibo almacen
    SQL = SQL & ", codcarpetareccamp = " & DBSet(mvarCarpetaRecCampo, "N") 'carpeta recibo campo
    SQL = SQL & ", codextension = " & DBSet(mvarExtension, "N") 'extension documento
    SQL = SQL & ", c1albaridoc = " & DBSet(mvarC1Albaran, "N") 'campo 1 albaran
    SQL = SQL & ", c2albaridoc = " & DBSet(mvarC2Albaran, "N") 'campo 2 albaran
    SQL = SQL & ", c3albaridoc = " & DBSet(mvarC3Albaran, "N") 'campo 3 albaran
    SQL = SQL & ", c4albaridoc = " & DBSet(mvarC4Albaran, "N") 'campo 4 albaran
    SQL = SQL & ", c1facaridoc = " & DBSet(mvarC1Factura, "N") 'campo 1 factura
    SQL = SQL & ", c2facaridoc = " & DBSet(mvarC2Factura, "N") 'campo 2 factura
    SQL = SQL & ", c3facaridoc = " & DBSet(mvarC3Factura, "N") 'campo 3 factura
    SQL = SQL & ", c4facaridoc = " & DBSet(mvarC4Factura, "N") 'campo 4 factura
    SQL = SQL & ", c1recaridoc = " & DBSet(mvarC1Recibo, "N") 'campo 1 recibo
    SQL = SQL & ", c2recaridoc = " & DBSet(mvarC2Recibo, "N") 'campo 2 recibo
    SQL = SQL & ", c3recaridoc = " & DBSet(mvarC3Recibo, "N") 'campo 3 recibo
    SQL = SQL & ", c4recaridoc = " & DBSet(mvarC4Recibo, "N") 'campo 4 recibo
    
    'integracion edi
    SQL = SQL & ", codigoedi= " & DBSet(mvarCodigoEdi, "T") ' codigo edi del vendedor
    SQL = SQL & ", regmercantil= " & DBSet(mvarRegMercantil, "T") ' registro mercantil
    SQL = SQL & ", pathedicom= " & DBSet(Replace(mvarPathEdicom, "\", "\\"), "T") ' path ficheros para edicom
    
    SQL = SQL & ", porcdesvia= " & DBSet(mvarPorcDesv, "N") ' porcentaje de desviacion para el calculo de gastos costes reales
    SQL = SQL & ", porteskilocaja= " & DBSet(mvarPortesKiloCaja, "N") ' calculo de portes por kilo o por caja
        
    SQL = SQL & ", codtipomalb= " & DBSet(mvarCodTipomAlb, "T") ' codtipom de albaranes de venta
        
    ' textos cmr
    SQL = SQL & ", text1cmr13= " & DBSet(mvarText1CMR13, "T", "S") ' texto 1 de cmr seccion 13
    SQL = SQL & ", text2cmr13= " & DBSet(mvarText2CMR13, "T", "S") ' texto 2 de cmr seccion 13
    
    SQL = SQL & ", codclien= " & DBSet(mvarClienteVtas, "N", "S") ' cliente de ventas
    SQL = SQL & ", calculocomision = " & DBSet(mvarTipoCalculoComision, "N") 'Tipo de calculo de comision en facturas de cliente
        
'    SQL = SQL & " WHERE codparam =" & Codigo & ";"
    SQL = SQL & ", hayccostes= " & DBSet(mvarHayCCostes, "N") ' hay control de costes
    
    '[Monica]11/01/2012: arigesmail
    SQL = SQL & ", EnvioDesdeOutlook = " & DBSet(Me.EnvioDesdeOutlook, "B")
    SQL = SQL & ", arigesmail = " & DBSet(ExeEnvioMail, "T", "S")

    '[Monica]25/05/2012: Path de fichadas de costes
    SQL = SQL & ", pathfichadas= " & DBSet(Replace(mvarPathFichadas, "\", "\\"), "T") ' path ficheros para fichadas para costes

    'FacturaE
    SQL = SQL & ", pathfacturae = " & DBSet(Replace(mvarPathFacturaE, "\", "\\"), "T")

    ' Integracion Anecoop
    SQL = SQL & ", hayanecoop= " & DBSet(mvarHayAnecoop, "N") ' hay anecoop

    conn.Execute SQL

    Modificar = 1
    Exit Function

EModificar:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Modificar = 0
End Function

