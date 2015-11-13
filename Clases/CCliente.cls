VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CLIENTES
'TABLA: sclien


'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodClien As String 'Cod. cliente
Private mNomClien As String 'Nombre cliente
Private mDomClien As String 'Domicilio cliente
Private mCodpobla As String 'cod. postal cliente
Private mPobClien As String 'poblacion cliente
Private mProClien As String 'Provincia cliente
Private mNIFClien As String 'NIF cliente
Private mTfnClien As String 'Telefono administracion
Private mMaiClien As String 'Mail administracion

Private mCodEnvio As String 'cod. envio
Private mCodAgent As String 'cod. agente
Private mCodForpa As String 'cod. forma de pago
Private mCodPais As String 'cod. pais

Private mTipoDto As Byte   'tipo dto (0:aditivo, 1:normal)
Private mDto1 As String 'descuento 1
Private mDto2 As String 'descuento 2
Private mTipoFact As Byte   'tipo facturacion (0:factura colectiva, 1:facturaxAlbaran
Private mTipoIva As Byte   'tipo iva (0:normal, 1:exento)

'[Monica]22/11/2013: tema iban
Private mIban As String ' iban

Private mCodbanco As String 'cod. banco
Private mCodsucur As String 'cod. sucur
Private mDigContr As String 'digito control
Private mCuentaba As String 'cuenta bancaria

Private mObservac As String 'Observaciones cliente

Private mWwwClien As String 'web del cliente

Private mTipoMov As String ' tipo de movimiento del cliente
Private mCodigoEdi As String ' codigo edi
Private mDestEDI As Byte   'destino de factura edicom (0:cliente, 1:destino)


'------------------------------------------------
'Propiedades del modulo CCliente
'------------------------------------------------

'**** Codigo del cliente

Public Property Let Codigo(ByVal vData As String)
     mCodClien = vData
End Property

Public Property Get Codigo() As String
     Codigo = mCodClien
End Property


'**** Nombre del cliente

Public Property Let Nombre(ByVal vData As String)
     mNomClien = vData
End Property

Public Property Get Nombre() As String
     Nombre = mNomClien
End Property


'**** Domicilio del cliente

Public Property Let Domicilio(ByVal vData As String)
     mDomClien = vData
End Property

Public Property Get Domicilio() As String
     Domicilio = mDomClien
End Property

'**** CPostal del cliente

Public Property Let CPostal(ByVal vData As String)
     mCodpobla = vData
End Property

Public Property Get CPostal() As String
     CPostal = mCodpobla
End Property

'**** Poblacion del cliente

Public Property Let Poblacion(ByVal vData As String)
     mPobClien = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mPobClien
End Property

'**** Provincia del cliente

Public Property Let Provincia(ByVal vData As String)
     mProClien = vData
End Property

Public Property Get Provincia() As String
     Provincia = mProClien
End Property

'**** NIF del cliente

Public Property Let NIF(ByVal vData As String)
     mNIFClien = vData
End Property

Public Property Get NIF() As String
     NIF = mNIFClien
End Property

'**** Telefono administracion del cliente

Public Property Let TfnoClien(ByVal vData As String)
     mTfnClien = vData
End Property

Public Property Get TfnoClien() As String
     TfnoClien = mTfnClien
End Property

'**** e-MAil administracion del cliente

Public Property Let EMailAdm(ByVal vData As String)
     mMaiClien = vData
End Property

Public Property Get EMailAdm() As String
     EMailAdm = mMaiClien
End Property


'**** Web del cliente

Public Property Let WebClien(ByVal vData As String)
     mWwwClien = vData
End Property

Public Property Get WebClien() As String
     WebClien = mWwwClien
End Property

'**** Tipo Dto del cliente

Public Property Let TipoDto(ByVal vData As Byte)
     mTipoDto = vData
End Property

Public Property Get TipoDto() As Byte
     TipoDto = mTipoDto
End Property


'**** Tipo Facturacion del cliente

Public Property Let TipoFactu(ByVal vData As Byte)
     mTipoFact = vData
End Property

Public Property Get TipoFactu() As Byte
     TipoFactu = mTipoFact
End Property

'**** Tipo Iva del cliente

Public Property Let TipoIva(ByVal vData As Byte)
     mTipoIva = vData
End Property

Public Property Get TipoIva() As Byte
     TipoIva = mTipoIva
End Property

'**** Forma de pago del Cliente

Public Property Let ForPago(ByVal vData As String)
     mCodForpa = vData
End Property

Public Property Get ForPago() As String
     ForPago = mCodForpa
End Property

'**** Código del país del Cliente

Public Property Let CodPais(ByVal vData As String)
     mCodPais = vData
End Property

Public Property Get CodPais() As String
     CodPais = mCodPais
End Property




'**** Descuento 1

Public Property Let Dto1(ByVal vData As String)
     mDto1 = vData
End Property

Public Property Get Dto1() As String
     Dto1 = mDto1
End Property

'**** Descuento 2

Public Property Let Dto2(ByVal vData As String)
     mDto2 = vData
End Property

Public Property Get Dto2() As String
     Dto2 = mDto2
End Property


'**** Banco del Cliente

Public Property Let Banco(ByVal vData As String)
     mCodbanco = vData
End Property

Public Property Get Banco() As String
     Banco = mCodbanco
End Property


'**** Sucursal del Cliente

Public Property Let Sucursal(ByVal vData As String)
     mCodsucur = vData
End Property

Public Property Get Sucursal() As String
     Sucursal = mCodsucur
End Property


'**** Digito Control de la cuenta del cliente

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


'[Monica]22/11/2013: tema Iban
'**** IBAN
Public Property Let Iban(ByVal vData As String)
     mIban = vData
End Property

Public Property Get Iban() As String
     Iban = mIban
End Property


'**** Observaciones del cliente

Public Property Let Observaciones(ByVal vData As String)
     mObservac = vData
End Property

Public Property Get Observaciones() As String
     Observaciones = mObservac
End Property

'**** Codtipom del cliente

Public Property Let tipoMov(ByVal vData As String)
     mTipoMov = vData
End Property

Public Property Get tipoMov() As String
     tipoMov = mTipoMov
End Property

'**** Codigo edi del cliente

Public Property Let CodigoEdi(ByVal vData As String)
     mCodigoEdi = vData
End Property

Public Property Get CodigoEdi() As String
     CodigoEdi = mCodigoEdi
End Property

'**** Destino de Factura EDICOM

Public Property Let DestEDI(ByVal vData As Byte)
     mDestEDI = vData
End Property

Public Property Get DestEDI() As Byte
     DestEDI = mDestEDI
End Property



'------------------------------------------------
'Procedimientos del modulo CCliente
'------------------------------------------------

Public Function Existe(vCodClien As String) As Boolean
'Comprueba si existe el cliente en la BD
Dim devuelve As String

    On Error GoTo EExiste
    
    devuelve = DevuelveDesdeBDNew(cAgro, "clientes", "codclien", "codclien", vCodClien, "N")
    If devuelve = "" Then
        Existe = False
        devuelve = "No existe el Cliente " & vCodClien & ". "
        MsgBox devuelve, vbExclamation
    Else
        Existe = True
    End If
    
EExiste:
    If Err.Number <> 0 Then Existe = False
End Function



Public Function LeerDatos(vCodClien As String) As Boolean
'Leer los datos de un cliente dado
'Lee de la BD: Ariges, Tabla: sclien
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ELeer
    
    LeerDatos = False
 'codclien,nomclien,domclien,codpobla,pobclien,proclien,codpaise,wwwclien,cifclien,telclie1,faxclie1,
 'movclie1,maiclie1,perclie1,telclie2,faxclie2,movclie2,maiclie2,perclie2,codforpa,tipoiva,nrosegur,
 'diassegu,limcredi,porccom1,porccom2,tipodtos,codmacta,cliabono,tipofact,codtipom
 
    SQL = "SELECT codclien,nomclien,codpobla,domclien,pobclien,proclien,codpaise,cifclien, "
    SQL = SQL & "wwwclien,codforpa,codbanco,codsucur,digcontr,cuentaba,"
    SQL = SQL & "telclie1,maiclie1,porccom1,porccom2,codmacta,tipofact,codtipom, "
    SQL = SQL & "observac,tipoiva, tipodtos, codigoedi, destedi, codpaise "
    
    '[Monica]22/11/2013: tema iban
    SQL = SQL & ", iban "
        
    SQL = SQL & " FROM clientes "
    SQL = SQL & " WHERE codclien=" & vCodClien
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If RS.EOF Then
        LeerDatos = False
    Else
        mCodClien = CStr(RS!CodClien)
        
        mNomClien = RS!nomclien
        mDomClien = DBLet(RS!domclien, "T")
        mCodpobla = DBLet(RS!codPobla, "T")
        mPobClien = DBLet(RS!pobclien, "T")
        mProClien = DBLet(RS!proclien, "T")
        mNIFClien = DBLet(RS!cifClien, "T")
        mWwwClien = DBLet(RS!wwwclien, "T")
        mTfnClien = DBLet(RS!telclie1, "T")
        mMaiClien = DBLet(RS!maiclie1, "T")
       
        mCodForpa = CStr(RS!Codforpa)
        mCodPais = CStr(DBLet(RS!CodPaise, "N"))
        
        mDto1 = CStr(DBLet(RS!porccom1, "N"))
        mDto2 = CStr(DBLet(RS!porccom2, "N"))
        mTipoFact = RS!TipoFact
        mTipoIva = DBLet(RS!TipoIva, "N")
        mTipoDto = DBLet(RS!TipoDtos, "N")
        
        
        '[Monica]22/11/2013: tema iban
        mIban = CStr(DBLet(RS!Iban, "T"))
        
        
        mCodbanco = CStr(DBLet(RS!codbanco, "N"))
        '[Monica]21/11/2014: le damos formato
        If mCodbanco <> "" And ComprobarCero(mCodbanco) <> 0 Then mCodbanco = Format(CInt(mCodbanco), "0000")
        
        mCodsucur = CStr(DBLet(RS!codsucur, "N"))
        '[Monica]21/11/2014: le damos formato
        If mCodsucur <> "" And ComprobarCero(mCodsucur) <> 0 Then mCodsucur = Format(CInt(mCodsucur), "0000")
        
        mDigContr = DBLet(RS!digcontr, "T")
        mCuentaba = DBLet(RS!cuentaba, "T")
        
        mObservac = DBLet(RS!Observac, "T")
        
        mTipoMov = DBLet(RS!codTipoM, "T")
        mCodigoEdi = DBLet(RS!CodigoEdi, "T")
        mDestEDI = DBLet(RS!DestEDI, "N")
        
        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Datos Cliente" & vbCrLf
    SQL = SQL & "Número: " & Err.Number & vbCrLf
    SQL = SQL & "Descripción: " & Err.Description
    MsgBox SQL, vbExclamation
    Set RS = Nothing
    LeerDatos = False
End Function



Public Function LeerDatosCliVario(vNIF As String) As Boolean
'si es cliente de varios recuperamos los datos nombre,direc,etc.
'de la tabla de varios: sclvar
Dim RS As ADODB.Recordset
Dim SQL As String
    
    On Error GoTo ELeer
    LeerDatosCliVario = False
    
    SQL = "SELECT nifclien,nomclien,domclien,codpobla,pobclien,proclien,telclien "
    SQL = SQL & " FROM sclvar "
    SQL = SQL & " WHERE nifclien=" & DBSet(vNIF, "T")
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If RS.EOF Then
        LeerDatosCliVario = False
        
        mNomClien = ""
        mDomClien = ""
        mCodpobla = ""
        mPobClien = ""
        mProClien = ""
        mTfnClien = ""
        MsgBox "No existe el cliente vario: " & vNIF, vbInformation
    Else
        mNIFClien = RS!nifClien
        mNomClien = RS!nomclien
        mDomClien = RS!domclien
        mCodpobla = DBLet(RS!codPobla, "T")
        mPobClien = DBLet(RS!pobclien, "T")
        mProClien = DBLet(RS!proclien, "T")
        mTfnClien = DBLet(RS!telclien, "T")
        
        LeerDatosCliVario = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Datos Cliente de Varios." & vbCrLf
    SQL = SQL & "Número: " & Err.Number & vbCrLf
    SQL = SQL & "Descripción: " & Err.Description
    MsgBox SQL, vbExclamation
    Set RS = Nothing
    LeerDatosCliVario = False
End Function



Public Function ActualizarClienteV(vNIF As String) As Boolean
'Actualiza los datos del cliente de varios en la tabla: sclvar
Dim SQL As String
Dim devuelve As String

    On Error GoTo EActualizarCliente
    ActualizarClienteV = False
    
    devuelve = DevuelveDesdeBDNew(cAgro, "sclvar", "nifclien", "nifclien", vNIF, "T")
    
    If devuelve <> "" Then 'Existe el cliente de varios y lo actualizamos
        SQL = "UPDATE sclvar SET nomclien=" & DBSet(Nombre, "T") & ", domclien=" & DBSet(Domicilio, "T") & ", "
        SQL = SQL & " codpobla=" & DBSet(CPostal, "T") & ", pobclien=" & DBSet(Poblacion, "T") & ", "
        SQL = SQL & " proclien=" & DBSet(Provincia, "T") & ", telclien=" & DBSet(TfnoClien, "T")
        SQL = SQL & " WHERE nifclien=" & DBSet(vNIF, "T")
    Else 'No existe el cliente y lo añadimos
        SQL = "INSERT INTO sclvar (nifclien, nomclien, domclien, codpobla, pobclien, proclien, telclien) "
        SQL = SQL & " VALUES (" & DBSet(vNIF, "T") & ", " & DBSet(Nombre, "T") & ", " & DBSet(Domicilio, "T") & ", "
        SQL = SQL & DBSet(CPostal, "T") & ", " & DBSet(Poblacion, "T") & ", " & DBSet(Provincia, "T") & ", " & DBSet(TfnoClien, "T") & ")"
    End If
    conn.Execute SQL, , adCmdText
    
    ActualizarClienteV = True
    
EActualizarCliente:
    If Err.Number <> 0 Then
        ActualizarClienteV = False
        MuestraError Err.Number, "Actualizando Cliente de Varios.", Err.Description
    End If
End Function

'--monica
'Public Function ClienteBloqueado() As Boolean
''devuelve true si el cliente esta bloqueado
''si la situación del cliente es distinta de NORMAL(codsitua=0) entonces
''mostrar un mensaje con la situación especial del cliente
'Dim Tipo As String
'Dim devuelve As String
'
'    On Error GoTo EBloqueado
'    ClienteBloqueado = False
'
'    If Situacion <> 0 Then '0: situacion normal
'        Tipo = "tipositu"
'        devuelve = DevuelveDesdeBDNew(cAgro, "ssitua", "nomsitua", "codsitua", Situacion, "N", Tipo)
'
'        If Tipo = "1" Then 'Cliente Bloqueado por Situación Especial.
'            MsgBox UCase("Cliente Bloqueado por: ") & vbCrLf & devuelve, vbInformation, "Situación Especial del Cliente."
'            ClienteBloqueado = True
'        Else
'            MsgBox devuelve, vbInformation, "Situación Especial del Cliente."
'        End If
'    End If
'
'EBloqueado:
'    If Err.Number <> 0 Then Err.Clear
'End Function



Public Function DptoCliente(vCodDpto As String, NomDpto As String) As Boolean
'comprueba si existe el departamento para el cliente y devuelve su nombre
'IN: vCodDpto= codigo del departamento o direccion del cliente
'OUT: NomDppto= devuelve el nombre del dpto o direccion si lo encuentra
Dim devuelve As String
Dim cadDpto As String
    
    If Codigo = "" Then
        devuelve = "Debe introducir valor al campo Cliente, para cargar sus "
'--monica
'        If vParamAplic.Departamento Then 'Se trabaja con dpto
'            devuelve = devuelve & "Departamentos. "
'        Else 'la aplicacion trabaja con direcciones
            devuelve = devuelve & "Direcciones. "
'        End If
        MsgBox devuelve, vbInformation
        Exit Function
    End If

    'Comprobar que el cliente seleccionado tiene esa direccion
    devuelve = DevuelveDesdeBDNew(cAgro, "sdirec", "nomdirec", "codclien", Codigo, "N", , "coddirec", vCodDpto, "N")
    NomDpto = devuelve
    If devuelve = "" Then
        DptoCliente = False
'--monica
'        If vParamAplic.Departamento Then 'Se trabaja con dpto
'            cadDpto = " el Departamento "
'        Else 'la aplicacion trabaja con direcciones
            cadDpto = " la Dirección "
'        End If
        
        devuelve = "No existe" & cadDpto & vCodDpto & " para el cliente: "
        devuelve = devuelve & Codigo
        MsgBox devuelve, vbInformation
    Else
        DptoCliente = True
    End If
End Function


Public Function TieneRefObligatoria(vReferen As String) As Boolean
'Comprueba si el campo de referencia obligatoria es necesario que se rellene
'al introducir datos de ese cliente
Dim devuelve As String

    devuelve = DevuelveDesdeBDNew(cAgro, "sclien", "referobl", "codclien", Codigo, "N")
    If devuelve = "1" And vReferen = "" Then 'Referencia Obligatoria
        TieneRefObligatoria = True
        devuelve = "La Referencia del Cliente es Obligatoria."
        MsgBox devuelve, vbInformation
    Else
        TieneRefObligatoria = False
    End If
End Function


Public Function exentoIVA() As Boolean
'Comprueba si el cliente esta exento de IVA para calcular el %IVA de las ofertas,pedidos
'albaranes, facturas,...
Dim devuelve As String

    devuelve = DevuelveDesdeBDNew(cAgro, "sclien", "tipoiva", "codclien", Codigo, "N")
    If devuelve = "2" Then 'exento de iva
        exentoIVA = True
    Else
        exentoIVA = False
    End If
End Function


Public Function Bonifica() As Boolean
'Indica si al cliente se le aplican bonificaciones o no
Dim cad As String

    On Error GoTo EBonifica
    Bonifica = False
    
    cad = DevuelveDesdeBDNew(cAgro, "sclien", "codtarif", "codclien", Codigo, "N")
    If cad <> "" Then
        cad = DevuelveDesdeBDNew(cAgro, "starif", "bonifica", "codlista", cad, "N")
        If cad <> "" Then Bonifica = CBool(cad)
    End If
    
EBonifica:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar bonificación del cliente.", Err.Description
    End If
End Function

'--monica
'Public Function ActualizaUltFecMovim(vFecha As String) As Boolean
''Actualiza la fecha del ultimo movimiento del cliente si la fecha del movimiento
''que se pasa por parametro es posterior a la que tiene el cliente
'Dim SQL As String
'Dim b As Boolean
'Dim actualiza As Boolean
'
'    On Error GoTo EActFecha
'
'    If FechaUltMov = "" Then _
'        FechaUltMov = DevuelveDesdeBDNew(cAgro, "sclien", "fechamov", "codclien", Codigo, "N")
'
'
'    If FechaUltMov = "" Then
'        actualiza = True
'    ElseIf CDate(vFecha) > CDate(FechaUltMov) Then
'        actualiza = True
'    Else
'        actualiza = False
'    End If
'
'    If actualiza Then
'        SQL = "UPDATE sclien SET fechamov=" & DBSet(vFecha, "F")
'        SQL = SQL & " WHERE codclien=" & Codigo
'        Conn.Execute SQL, , adCmdText
'    End If
'    b = True
'
'EActFecha:
'    If Err.Number <> 0 Then b = False
'    ActualizaUltFecMovim = b
'End Function
'



Public Sub MostrarObservaciones()
'Recuperamos el texto de las observaciones del cliente, y si tiene
'valor lo mostramos en un mensaje
    If Me.Observaciones <> "" Then _
         MsgBox Me.Observaciones, vbInformation, "Observaciones del cliente"
End Sub




Public Function ComprobarCtaBancaria(cadMen As String) As Boolean
'comprueba q el cliente tenga cuenta bancaria y que sea correcta
Dim cta As String

    If Me.Banco = 0 And Me.Sucursal = 0 And Me.DigControl = "" And Me.CuentaBan = "" Then
        ComprobarCtaBancaria = False
        cadMen = "El cliente no tiene asignada cuenta bancaria."
        Exit Function
    End If
    
    
    
    cta = Format(Me.Banco, "0000") & Format(Me.Sucursal, "0000") & Format(Me.DigControl, "00") & Format(Me.CuentaBan, "0000000000")
    
    If Val(ComprobarCero(cta)) = 0 Then
        ComprobarCtaBancaria = False
        cadMen = "El cliente no tiene asignada cuenta bancaria."
        Exit Function
    End If
    
'--monica
'    If Not Comprueba_CC(cta) Then
'        cadMen = "La cuenta bancaria del cliente no es correcta."
'        ComprobarCtaBancaria = False
'    Else
        ComprobarCtaBancaria = True
'    End If
End Function
