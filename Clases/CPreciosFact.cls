VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPreciosFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Modulo para Calcula el Precio a Aplicar a las lineas de una Oferta, Pedido o Albaran
' Y los descuentos correspondientes si se aplica descuento
'TABLAS: slista, sprees, sdtofm

'------------------------------------------------
'Variables del modulo CPreciosFact
'------------------------------------------------

Private mvarDtoPermitido As Boolean
Private mvarCodArtic As String
Private mvarCodLista As String
Private mvarCodClien As String
Private mvarCodFamia As String
Private mvarCodMarca As String

Private mvarDescuento1 As String
Private mvardescuento2 As String


'------------------------------------------------
'Propiedades del modulo CPreciosFact
'------------------------------------------------

Public Property Let DtoPermitido(ByVal vData As Boolean)
     mvarDtoPermitido = vData
End Property

Public Property Get DtoPermitido() As Boolean
     DtoPermitido = mvarDtoPermitido
End Property

Public Property Let CodigoArtic(ByVal vData As String)
     mvarCodArtic = vData
End Property

Public Property Get CodigoArtic() As String
     CodigoArtic = mvarCodArtic
End Property

Public Property Let CodigoLista(ByVal vData As String)
     mvarCodLista = vData
End Property

Public Property Get CodigoLista() As String
     CodigoLista = mvarCodLista
End Property


Public Property Let CodigoClien(ByVal vData As String)
     mvarCodClien = vData
End Property

Public Property Get CodigoClien() As String
     CodigoClien = mvarCodClien
End Property

Public Property Let CodigoFamia(ByVal vData As String)
     mvarCodFamia = vData
End Property

Public Property Get CodigoFamia() As String
     CodigoFamia = mvarCodFamia
End Property

Public Property Let CodigoMarca(ByVal vData As String)
     mvarCodMarca = vData
End Property

Public Property Get CodigoMarca() As String
     CodigoMarca = mvarCodMarca
End Property

Public Property Let Descuento1(ByVal vData As String)
     mvarDescuento1 = vData
End Property

Public Property Get Descuento1() As String
     Descuento1 = mvarDescuento1
End Property


Public Property Let Descuento2(ByVal vData As String)
     mvardescuento2 = vData
End Property

Public Property Get Descuento2() As String
     Descuento2 = mvardescuento2
End Property

'------------------------------------------------
'Metodos del modulo CPreciosFact
'------------------------------------------------

Public Function ObtenerPrecio(PCaja As Boolean, FechaOfe As String, OrigP As String) As String
'PCaja -> IN: si vale true se obiene el precio de unidad si se vende por caja sino se obtiene precio
'         OUT: si sale con precio de pvp del articulo se pone a False, independientemente de que se
'              pueda vender por cajas o no
'FechaOfe -> IN: Fecha de la Oferta para comprobar si la oferta esta en los periodos de promocion, etc...
'OrigP -> OUT: Aqui se devolvera con que valor precio se sale de la función, e.d.:
'               P=sale con precio de Promoción
'               E=sale con Precio Especial
'               etc....
'RETURN: valor del precio
Dim PrecioFact As String
Dim Dto1Fact As String
Dim Dto2Fact As String
On Error GoTo EObtenerPre
   
    'COMPROBAR PROMOCIONES
    '==============================================================
    'Buscar Precio en la tabla spromo si encuentro valor salgo con ese precio
    PrecioFact = ObtenerPromocion(FechaOfe, PCaja)
    If PrecioFact <> "" Then 'Ha obtenido el Precio
        If DtoPermitido Then
        'Buscar en la tabla sdtofm los descuentos
            Dto1Fact = ObtenerDescuentos(FechaOfe, Dto2Fact, PCaja)
        End If
        OrigP = "P" 'Indica que sale con precio de Promocion
        
    Else 'No encuentra Precio en spromo y Buscar en Sprees(Tabla Precios Especiales)
        'COMPROBAR PRECIO ESPECIAL
        '===========================================================================
        PrecioFact = ObtenerPrecioEsp(FechaOfe, PCaja, Dto1Fact)
        Dto2Fact = "0"
        If PrecioFact <> "" Then 'Ha obtenido el precio
            If DtoPermitido And Dto1Fact = "" Then
            'Buscar en la tabla sdtofm los descuentos
                Dto1Fact = ObtenerDescuentos(FechaOfe, Dto2Fact, PCaja)
            End If
            OrigP = "E" 'Indica que sale con Precio Especial
        Else 'No encuentra precio en sprees y Buscar en slista(Tabla Tarifas Articulos)
            'COMPROBAR TARIFAS ARTICULOS
            '=======================================================================
            PrecioFact = ObtenerTarifaArt(FechaOfe, PCaja)
            If PrecioFact <> "" Then 'Ha obtenido el precio
                If DtoPermitido Then
                    'Buscar en la tabla sdtofm los descuentos
                    Dto1Fact = ObtenerDescuentos(FechaOfe, Dto2Fact, PCaja)
                End If
                OrigP = "T"
            Else 'Buscar en sartic el precio del articulo (preciove)
                PrecioFact = ObtenerPrecioVenta
                PCaja = False
                OrigP = "A"
            End If
        End If
    End If
    ObtenerPrecio = ComprobarCero(PrecioFact)
    Descuento1 = ComprobarCero(Dto1Fact)
    Descuento2 = ComprobarCero(Dto2Fact)
    
EObtenerPre:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obtener Precio", Err.Description
End Function



Private Function ObtenerPromocion(FechaOfe As String, PorCaja As Boolean) As String
'Comprueba si existe una promocion en la tabla: spromo (Promociones Tarifas)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Precio As String
On Error GoTo EObtenerPromo

    Precio = ""
    
    'Obtener Precio Actual
    SQL = "SELECT * FROM spromo "
    SQL = SQL & " WHERE codartic=" & DBSet(CodigoArtic, "T") & " AND codlista=" & CodigoLista
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
    'Existe una promocion para ese articulo y tarifa, comprobar fechas
        If EntreFechas(DBLet(RS!FechaIni), FechaOfe, DBLet(RS!FechaFin)) Then
        'Dentro Periodo Promocion Actual
            If PorCaja Then
                Precio = ComprobarCero(DBLet(RS!precioa1))
            Else
                Precio = DBLet(RS!precioac)
            End If
        ElseIf DBLet(RS!fechain1) = "" And DBLet(RS!fechafi1) = "" Then
            Precio = ""
        ElseIf EntreFechas(DBLet(RS!fechain1), FechaOfe, DBLet(RS!fechafi1)) Then
        'DEntro Periodo Promocion Nueva
            If PorCaja Then
                Precio = ComprobarCero(DBLet(RS!precion1))
            Else
                Precio = ComprobarCero(DBLet(RS!precionu))
            End If
        End If
        DtoPermitido = CBool(RS!dtopermi)
    Else 'No existe Promocion salir y pasar al siguiente paso
        Precio = ""
        DtoPermitido = False
    End If
    RS.Close
    Set RS = Nothing
    ObtenerPromocion = Precio
    
EObtenerPromo:
    If Err.Number <> 0 Then MuestraError Err.Number, "Precio Promoción", Err.Description
End Function

Private Function ObtenerDescuentos(FechaOfe As String, Dto2Fact, PorCaja As Boolean) As String
Dim RS As ADODB.Recordset
Dim SQL As String
Dim CampoDto1 As String, campoDto2 As String
Dim vCodMarca As String

    vCodMarca = "Codmarca"
    CodigoFamia = DevuelveDesdeBDNew(cAgro, "sartic", "Codfamia", "codartic", CodigoArtic, "T", vCodMarca)
    CodigoMarca = vCodMarca
    If PorCaja Then 'Obtener Descuentos Por Caja
        CampoDto1 = "dtocaja1"
        campoDto2 = "dtocaja2"
    Else 'Obtener Descuentos por Unidad
        CampoDto1 = "dtoline1"
        campoDto2 = "dtoline2"
    End If

    'Obtener Descuento
    SQL = "SELECT " & CampoDto1 & ", " & campoDto2 & ", codlista FROM sdtofm "
    SQL = SQL & " WHERE codclien=" & CodigoClien
    SQL = SQL & " AND (fechadto<= '" & Format(FechaOfe, FormatoFecha) & "')"
    SQL = SQL & " AND ((codfamia=" & CodigoFamia & " AND codmarca=" & CodigoMarca & ")"
    SQL = SQL & " OR (codfamia=" & CodigoFamia & " AND codmarca=0)"
    SQL = SQL & " OR (codfamia=0 AND codmarca=" & CodigoMarca & ")"
    SQL = SQL & " OR (codfamia=0 AND codmarca=0))"
'**quitado
'    If vParamAplic.OrdenDtos = 0 Then
'        SQL = SQL & " ORDER BY codfamia DESC, codmarca DESC"
'    Else
'        SQL = SQL & " ORDER BY codmarca DESC,codfamia DESC"
'    End If
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
    'Existe el descuento, y salimos con los descuentos
        ObtenerDescuentos = DBLet(RS.Fields(0).Value)
        Dto2Fact = DBLet(RS.Fields(1).Value)
    Else
        'Buscar si existe en Promociones Nueva
        ObtenerDescuentos = ""
        Dto2Fact = ""
    End If
    RS.Close
    Set RS = Nothing
End Function



Private Function ObtenerPrecioEsp(FechaOfe As String, PorCaja As Boolean, Dto As String) As String
'Comprueba si existe un Precio Especial en la tabla: sprees (Precios Especiales)
'Dim devuelve As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Precio As String
On Error GoTo EObtenerPreEsp

    Precio = ""
    
    'Obtener Precio Actual
    SQL = "SELECT * FROM sprees "
    SQL = SQL & " WHERE codclien=" & CodigoClien & " AND codartic=" & DBSet(CodigoArtic, "T")

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
    'Existe una promocion, y salimos con el precio actual o nuevo
        DtoPermitido = CBool(RS!dtopermi)
        If IsNull(RS!fechanue) Or (EsFechaIgualPosterior(FechaOfe, DBLet(RS!fechanue), False)) Then
        'Precio Actual
            If PorCaja Then
                Precio = ComprobarCero(DBLet(RS!precioa1))
            Else
                Precio = DBLet(RS!precioac)
            End If
            If DtoPermitido Then Dto = DBLet(RS!dtoespec)
        ElseIf EsFechaIgualPosterior(RS!fechanue, FechaOfe, False) Then
        'Precio Nuevo
            If PorCaja Then
                Precio = ComprobarCero(DBLet(RS!precion1))
            Else
                Precio = ComprobarCero(DBLet(RS!precionu))
            End If
            If DtoPermitido Then Dto = DBLet(RS!dtoespe1)
        End If
    Else  'No existe Precio Especial salir y pasar al siguiente paso
        Precio = ""
        DtoPermitido = False
        Dto = ""
    End If
    RS.Close
    Set RS = Nothing
    ObtenerPrecioEsp = Precio
EObtenerPreEsp:
    If Err.Number <> 0 Then MuestraError Err.Number, "Precio Especia", Err.Description
End Function


Private Function ObtenerTarifaArt(FechaOfe As String, PorCaja As Boolean) As String
'Comprueba si existe una Tarifa de Articulos en la tabla: slista (Tarifas Articulos)
'Dim devuelve As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Precio As String

    On Error GoTo EObtenerTarifa

    Precio = ""
    
    'Obtener Precio Actual
    SQL = "SELECT * FROM slista "
    SQL = SQL & " WHERE codartic=" & DBSet(CodigoArtic, "T") & " AND codlista=" & CodigoLista
    'SQL = SQL & " and ( isnull(fechanue) or (fechanue<= '" & Format(FechaOfe, FormatoFecha) & "'))"

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
    'Existe una promocion, y salimos con el precio actual o nuevo
        DtoPermitido = CBool(RS!dtopermi)
        If IsNull(RS!fechanue) Or (EsFechaIgualPosterior(FechaOfe, DBLet(RS!fechanue), False)) Then
        'Precio Actual
            If PorCaja Then
                Precio = ComprobarCero(DBLet(RS!precioa1))
            Else
                Precio = DBLet(RS!precioac)
            End If
        ElseIf EsFechaIgualPosterior(RS!fechanue, FechaOfe, False) Then
        'Precio Nuevo
            If PorCaja Then
                Precio = ComprobarCero(DBLet(RS!precion1))
            Else
                Precio = ComprobarCero(DBLet(RS!precionu))
            End If
        End If
    Else  'No existe Precio Especial salir y pasar al siguiente paso
        Precio = ""
        DtoPermitido = False
    End If
    RS.Close
    Set RS = Nothing
    ObtenerTarifaArt = Precio
EObtenerTarifa:
    If Err.Number <> 0 Then MuestraError Err.Number, "Precio Tarifa Artículo", Err.Description
End Function


Private Function ObtenerPrecioVenta() As String
'Obtiene el precio de Venta al Publico (preciove) de la tabla sartic
    ObtenerPrecioVenta = DevuelveDesdeBDNew(cAgro, "sartic", "preciove", "codartic", CodigoArtic, "T")
End Function


Public Function ObtenerNumCajas(TUnidades As String, UniCaja As String) As Integer
Dim NumCajas As Integer
Dim Cantidad As Integer, UniPorCaja As Integer
On Error Resume Next

    Cantidad = CInt(TUnidades)
    UniPorCaja = CInt(UniCaja)
    If UniPorCaja > 1 Then 'Se vende en cajas
        NumCajas = Int(Cantidad / UniPorCaja)
    Else 'No se vende por cajas
        NumCajas = 0
    End If
    ObtenerNumCajas = NumCajas
End Function

