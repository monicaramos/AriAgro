Attribute VB_Name = "ModBasico"
Option Explicit


Public Sub AyudaFamilias(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Descripción|4095|;S|txtAux(2)|T|Cta.Ventas|1500|;S|txtAux(3)|T|Cta.Compras|1500|;"
    frmCom.CadenaConsulta = "SELECT sfamia.codfamia, sfamia.nomfamia, sfamia.ctaventa, sfamia.ctacompr "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sfamia "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999|sfamia|codfamia|000|S|"
    frmCom.Tag2 = "Descripción|T|N|||sfamia|nomfamia|||"
    frmCom.Tag3 = "Cta.Ventas|T|S|||sfamia|ctaventa|||"
    frmCom.Tag4 = "Cta.Compras|T|S|||sfamia|ctacompr|||"
    
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 25
    frmCom.Maxlen3 = 10
    frmCom.Maxlen4 = 10
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "sfamia"
    frmCom.CampoCP = "codfamia"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Familias"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 1000
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaPalets(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Nro.Palet|1105|;S|txtAux(1)|T|L.Confeccion|1500|;S|txtAux(2)|T|Tipo Palet|3095|;S|txtAux(3)|T|F.Inicio|1400|;S|txtAux(4)|T|F.Fin|1400|;"
    
    frmCom.CadenaConsulta = "SELECT palets.numpalet, palets.linconfe, confpale.nompalet, palets.fechaini, palets.fechafin "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM palets inner join confpale ON palets.codpalet=confpale.codpalet "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Nº Palet|N|S|||palets|numpalet|0000000|S|"
    frmCom.Tag2 = "Lin.Confe|N|N|||palets|linconfe|00||"
    frmCom.Tag3 = "Palet|T|S|||confpale|nompalet|||"
    frmCom.Tag4 = "F.Inicio|F|S|||palets|fechaini|dd/mm/yyyy||"
    frmCom.Tag5 = "F.Fin|F|S|||palets|fechafin|dd/mm/yyyy||"
    
    frmCom.Maxlen1 = 7
    frmCom.Maxlen2 = 3
    frmCom.Maxlen3 = 25
    frmCom.Maxlen4 = 10
    frmCom.Maxlen5 = 10
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "palets inner join confpale ON palets.codpalet=confpale.codpalet "
    frmCom.CampoCP = "palets.numpalet"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Palets"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 1500
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaAlbaranes(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmCom.CadenaTots = "S|txtAux(0)|T|Albaran|1105|;S|txtAux(1)|T|Cliente|1100|;S|txtAux(2)|T|Nombre|4495|;S|txtAux(3)|T|F.Albaran|1400|;"
    frmCom.CadenaConsulta = "SELECT albaran.numalbar, albaran.codclien, clientes.nomclien, albaran.fechaalb "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM albaran inner join clientes ON albaran.codclien=clientes.codclien "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Albaran|N|S|||albaran|numalbar|0000000|S|"
    frmCom.Tag2 = "Cliente|N|N|||albaran|codclien|000000||"
    frmCom.Tag3 = "Nombre|T|S|||clientes|nomclien|||"
    frmCom.Tag4 = "F.Albaran|F|S|||albaran|fecalbar|dd/mm/yyyy||"
    frmCom.Tag5 = ""
    
    frmCom.Maxlen1 = 7
    frmCom.Maxlen2 = 6
    frmCom.Maxlen3 = 25
    frmCom.Maxlen4 = 10
    frmCom.Maxlen5 = 0
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "albaran inner join clientes ON albaran.codclien=clientes.codclien "
    frmCom.CampoCP = "numalbar"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Albaranes de Clientes"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 1100 ' 1500
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaAlbaranesSocio(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmCom.CadenaTots = "S|txtAux(0)|T|Albaran|1105|;S|txtAux(1)|T|Socio|1100|;S|txtAux(2)|T|Nombre|4495|;S|txtAux(3)|T|F.Albaran|1400|;"
    frmCom.CadenaConsulta = "SELECT albaran.numalbar, albaran.codsocio, rsocios.nomsocio, albaran.fechaalb "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM albaran inner join rsocios ON albaran.codsocio=rsocios.codsocio "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Albaran|N|S|||albaran|numalbar|0000000|S|"
    frmCom.Tag2 = "Socio|N|N|||albaran|codsocio|000000||"
    frmCom.Tag3 = "Nombre|T|S|||rsocios|nomsocio|||"
    frmCom.Tag4 = "F.Albaran|F|S|||albaran|fecalbar|dd/mm/yyyy||"
    frmCom.Tag5 = ""
    
    frmCom.Maxlen1 = 7
    frmCom.Maxlen2 = 6
    frmCom.Maxlen3 = 25
    frmCom.Maxlen4 = 10
    frmCom.Maxlen5 = 0
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "albaran inner join rsocios ON albaran.codsocio=rsocios.codsocio "
    frmCom.CampoCP = "numalbar"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Albaranes Venta de Socios"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 1100 ' 1500
    
    frmCom.Show vbModal
End Sub



Public Sub AyudaPedidos(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Pedido|1105|;S|txtAux(1)|T|Cliente|1100|;S|txtAux(2)|T|Nombre|4495|;S|txtAux(3)|T|Ref.Cliente|1400|;S|txtAux(4)|T|Fec.Pedido|1400|;S|txtAux(5)|T|Fec.Carga|1400|;"
    
    frmCom.CadenaConsulta = "SELECT pedidos.numpedid, pedidos.codclien, clientes.nomclien, pedidos.refclien, pedidos.fechaped, pedidos.fechacar "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM pedidos inner join clientes ON pedidos.codclien=clientes.codclien "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Pedido|N|S|||pedidos|numpedid|0000000|S|"
    frmCom.Tag2 = "Cliente|N|N|||albaran|codclien|000000||"
    frmCom.Tag3 = "Nombre|T|S|||clientes|nomclien|||"
    frmCom.Tag4 = "Ref.Clien|T|S|||pedidos|refclien|||"
    frmCom.Tag5 = "F.Pedido|F|S|||pedidos|fechaped|dd/mm/yyyy||"
    frmCom.Tag6 = "F.Carga|F|S|||pedidos|fechacar|dd/mm/yyyy||"
    
    frmCom.Maxlen1 = 7
    frmCom.Maxlen2 = 6
    frmCom.Maxlen3 = 25
    frmCom.Maxlen4 = 10
    frmCom.Maxlen5 = 10
    frmCom.Maxlen6 = 10
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "pedidos inner join clientes ON pedidos.codclien=clientes.codclien "
    frmCom.CampoCP = "numpedid"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Pedidos"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 3900
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaFacturas(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmCom.CadenaTots = "S|txtAux(0)|T|Tipo|1105|;S|txtAux(1)|T|Factura|1100|;S|txtAux(2)|T|Fecha|1400|;S|txtAux(3)|T|Cliente|1100|;S|txtAux(4)|T|Nombre|4795|;"
    
    frmCom.CadenaConsulta = "SELECT facturas.codtipom, facturas.numfactu, facturas.fecfactu, facturas.codclien, clientes.nomclien "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM facturas inner join clientes ON facturas.codclien=clientes.codclien "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Tipo|T|N|||facturas|codtipom||S|"
    frmCom.Tag2 = "Factura|N|S|||facturas|numfactu|0000000|S|"
    frmCom.Tag3 = "F.Factura|F|S|||facturas|fecfactu|dd/mm/yyyy||"
    frmCom.Tag4 = "Cliente|N|N|||facturas|codclien|000000||"
    frmCom.Tag5 = "Nombre|T|S|||clientes|nomclien|||"
    
    frmCom.Maxlen1 = 5
    frmCom.Maxlen2 = 7
    frmCom.Maxlen3 = 10
    frmCom.Maxlen4 = 6
    frmCom.Maxlen5 = 30
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "facturas inner join clientes ON facturas.codclien=clientes.codclien "
    frmCom.CampoCP = "numfactu"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Facturas"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|2|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 2500
    
    frmCom.Show vbModal
End Sub



Public Sub AyudaFacturasTransporte(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmCom.CadenaTots = "S|txtAux(0)|T|Tipo|1505|;S|txtAux(1)|T|Factura|1300|;S|txtAux(2)|T|Fecha|1400|;S|txtAux(3)|T|Agencia|1100|;S|txtAux(4)|T|Nombre|4795|;"
    
    frmCom.CadenaConsulta = "SELECT if(tcafpc.tipo=0,'Transportista','Comisionista'), tcafpc.numfactu, tcafpc.fecfactu, tcafpc.codtrans, agencias.nomtrans "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM tcafpc inner join agencias ON tcafpc.codtrans=agencias.codtrans "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Tipo|N|N|0|1|tcafpc|tipo|||"
    frmCom.Tag2 = "Nº Factura|T|N|||tcafpc|numfactu||S|"
    frmCom.Tag3 = "Fecha Factura|F|N|||tcafpc|fecfactu|dd/mm/yyyy|S|"
    frmCom.Tag4 = "Agencia|N|N|0|999|tcafpc|codtrans|000|S|"
    frmCom.Tag5 = "Nombre|T|S|||agencias|nomtrans|||"
    
    frmCom.Maxlen1 = 5
    frmCom.Maxlen2 = 7
    frmCom.Maxlen3 = 10
    frmCom.Maxlen4 = 6
    frmCom.Maxlen5 = 30
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "tcafpc inner join agencias ON tcafpc.codtrans=agencias.codtrans "
    frmCom.CampoCP = "numfactu"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Facturas Transporte/Comisionista"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "1|2|3|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 3100
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaAlbaranEnvases(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmCom.CadenaTots = "S|txtAux(0)|T|Albaran|1105|;S|txtAux(1)|T|Cliente|1100|;S|txtAux(3)|T|Nombre|5195|;S|txtAux(2)|T|Fecha|1400|;"
    
    frmCom.CadenaConsulta = "SELECT scaalb.numalbar, scaalb.codclien, clientes.nomclien, scaalb.fechaalb "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scaalb inner join clientes ON scaalb.codclien=clientes.codclien "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "NºAlbarán|N|S|||scaalb|numalbar|0000000|S|"
    frmCom.Tag2 = "Cliente|N|N|0|999999|scaalb|codclien|000000||"
    frmCom.Tag3 = "Nombre|T|S|||clientes|nomclien|||"
    frmCom.Tag4 = "Fecha Albaran|F|N|||scaalb|fechaalb|dd/mm/yyyy|N|"
    
    frmCom.Maxlen1 = 7
    frmCom.Maxlen2 = 6
    frmCom.Maxlen3 = 35
    frmCom.Maxlen4 = 10
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "scaalb inner join clientes ON scaalb.codclien=clientes.codclien "
    frmCom.CampoCP = "numalbar"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Albaranes Venta Envases"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 1800
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaClientesPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmCom.CadenaTots = "S|txtAux(0)|T|Cliente|1100|;S|txtAux(1)|T|Nombre|4795|;S|txtAux(2)|T|NIF|1905|;S|txtAux(3)|T|Telefono|1900|;S|txtAux(3)|T|Móvil|1900|;"
    
    frmCom.CadenaConsulta = "SELECT clientes.codclien, clientes.nomclien, clientes.cifclien, clientes.telclie1, clientes.movclie1 "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM clientes  "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    
    frmCom.Tag1 = "Codigo|N|S|0|999999|clientes|codclien|000000|S|"
    frmCom.Tag2 = "Nombre|T|S|||clientes|nomclien|||"
    frmCom.Tag3 = "NIF|T|S|||clientes|cifclien|||"
    frmCom.Tag4 = "Teléfono|T|S|||clientes|telclie1|||"
    frmCom.Tag5 = "Móvil|T|S|||clientes|movclie1|||"
    
    frmCom.Maxlen1 = 6
    frmCom.Maxlen2 = 35
    frmCom.Maxlen3 = 15
    frmCom.Maxlen4 = 15
    frmCom.Maxlen5 = 15
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "clientes "
    frmCom.CampoCP = "codclien"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Clientes"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 4600
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaArticulosPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|1905|;S|txtAux(1)|T|Nombre|4195|;S|txtAux(2)|T|Codigo EAN|1900|;"
    frmCom.CadenaConsulta = "SELECT sartic.codartic, sartic.nomartic, sartic.codigoea"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|T|N|||sartic|codartic|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||sartic|nomartic|||"
    frmCom.Tag3 = "Código EAN|T|S|||sartic|codigoea|||"
    
    frmCom.Maxlen1 = 16
    frmCom.Maxlen2 = 40
    frmCom.Maxlen3 = 13
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "sartic"
    frmCom.CampoCP = "codartic"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Artículos"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 1000
    
    frmCom.Show vbModal
End Sub




Public Sub AyudaForfaits(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|1905|;S|txtAux(1)|T|Descripción|5095|;S|txtAux(2)|T|Variedad|1000|;S|txtAux(3)|T|Nombre Variedad|2500|;"
    frmCom.CadenaConsulta = "SELECT forfaits.codforfait, forfaits.nomconfe, forfaits.codvarie, variedades.nomvarie "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM forfaits left join variedades on forfaits.codvarie = variedades.codvarie "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código forfait|T|N|||forfaits|codforfait||S|"
    frmCom.Tag2 = "Nombre|T|N|||forfaits|nomconfe|||"
    frmCom.Tag3 = "Variedad|N|S|||forfaits|codvarie|000000||"
    frmCom.Tag4 = "Nombre Variedad|T|S|||variedades|nomvarie|||"
    
    frmCom.Maxlen1 = 16
    frmCom.Maxlen2 = 40
    frmCom.Maxlen3 = 6
    frmCom.Maxlen4 = 15
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "forfaits"
    frmCom.CampoCP = "codforfait"
    frmCom.TipoCP = "T"
    frmCom.Caption = "Forfaits"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 3530 '4400
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaProveedoresPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmCom.CadenaTots = "S|txtAux(0)|T|Proveedor|1100|;S|txtAux(1)|T|Nombre|4795|;S|txtAux(2)|T|NIF|1905|;S|txtAux(3)|T|Telefono|1900|;S|txtAux(3)|T|Fax|1900|;"
    
    frmCom.CadenaConsulta = "SELECT proveedor.codprove, proveedor.nomprove, proveedor.nifprove, proveedor.telprov1, proveedor.faxprov1 "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM proveedor  "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    
    frmCom.Tag1 = "Codigo|N|S|0|999999|proveedor|codprove|000000|S|"
    frmCom.Tag2 = "Nombre|T|S|||proveedor|nomprove|||"
    frmCom.Tag3 = "NIF|T|S|||proveedor|nifprove|||"
    frmCom.Tag4 = "Teléfono|T|S|||proveedor|telprov1|||"
    frmCom.Tag5 = "Móvil|T|S|||proveedor|faxprov1|||"
    
    frmCom.Maxlen1 = 6
    frmCom.Maxlen2 = 35
    frmCom.Maxlen3 = 15
    frmCom.Maxlen4 = 15
    frmCom.Maxlen5 = 15
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "proveedor"
    frmCom.CampoCP = "codprove"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Proveedores"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 4600
    
    frmCom.Show vbModal
End Sub



Public Sub AyudaPedidosCompraPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Pedido|1105|;S|txtAux(1)|T|Fecha|1400|;S|txtAux(2)|T|Código|1000|;S|txtAux(3)|T|Proveedor|4595|;"
    
    frmCom.CadenaConsulta = "SELECT scappr.numpedpr, scappr.fecpedpr, scappr.codprove, scappr.nomprove "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scappr "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Nº Pedido|N|S|0||scappr|numpedpr|0000000|S|"
    frmCom.Tag2 = "Fecha Pedido|F|N|||scappr|fecpedpr|dd/mm/yyyy|N|"
    frmCom.Tag3 = "Codigo|N|N|0|999999|scappr|codprove|000000|N|"
    frmCom.Tag4 = "Nombre Proveedor|T|N|||scappr|nomprove||N|"
    
    frmCom.Maxlen1 = 7
    frmCom.Maxlen2 = 10
    frmCom.Maxlen3 = 7
    frmCom.Maxlen4 = 35
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "scappr"
    frmCom.CampoCP = "numpedpr"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Pedidos Proveedores"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 1100
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaAlbaranesCompraPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Albaran|1105|;S|txtAux(1)|T|Fecha|1400|;S|txtAux(2)|T|Código|1000|;S|txtAux(3)|T|Proveedor|4595|;"
    
    frmCom.CadenaConsulta = "SELECT scaalp.numalbar, scaalp.fechaalb, scaalp.codprove, scaalp.nomprove "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scaalp "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Nº Albaran|N|S|0||scaalp|numalbar|0000000|S|"
    frmCom.Tag2 = "Fecha Albaran|F|N|||scaalp|fechaalb|dd/mm/yyyy|N|"
    frmCom.Tag3 = "Codigo|N|N|0|999999|scaalp|codprove|000000|N|"
    frmCom.Tag4 = "Nombre Proveedor|T|N|||scaalp|nomprove||N|"
    
    frmCom.Maxlen1 = 7
    frmCom.Maxlen2 = 10
    frmCom.Maxlen3 = 7
    frmCom.Maxlen4 = 35
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "scaalp"
    frmCom.CampoCP = "numalbar"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Albaranes Proveedores"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 1100
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaFacturasCompraPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Factura|1405|;S|txtAux(1)|T|Fecha|1400|;S|txtAux(2)|T|F.Recepción|1400|;S|txtAux(3)|T|Código|1000|;S|txtAux(4)|T|Proveedor|4795|;"
    
    frmCom.CadenaConsulta = "SELECT scafpc.numfactu, scafpc.fecfactu, scafpc.fecrecep, scafpc.codprove, scafpc.nomprove "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scafpc "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Nº Factura|T|S|||scafpc|numfactu||S|"
    frmCom.Tag2 = "Fecha Factura|F|N|||scafpc|fecfactu|dd/mm/yyyy|S|"
    frmCom.Tag3 = "Fecha Recepcion|F|N|||scafpc|fecrecep|dd/mm/yyyy|N|"
    frmCom.Tag4 = "Codigo|N|N|0|999999|scafpc|codprove|000000|S|"
    frmCom.Tag5 = "Nombre Proveedor|T|N|||scafpc|nomprove||N|"
    
    frmCom.Maxlen1 = 7
    frmCom.Maxlen2 = 10
    frmCom.Maxlen3 = 10
    frmCom.Maxlen4 = 7
    frmCom.Maxlen5 = 35
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "scafpc"
    frmCom.CampoCP = "numfactu"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Facturas Proveedores"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|3|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 3000
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaFacturasSociosPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Tipo|1400|;S|txtAux(1)|T|Factura|1405|;S|txtAux(2)|T|Fecha|1400|;S|txtAux(3)|T|Código|1000|;S|txtAux(4)|T|Socio|4795|;"
    
    frmCom.CadenaConsulta = "SELECT facturassocio.codtipom, facturassocio.numfactu, facturassocio.fecfactu, facturassocio.codsocio, rsocios.nomsocio "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM facturassocio inner join rsocios on facturassocio.codsocio = rsocios.codsocio "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Tipo Movimiento|T|N|||facturassocio|codtipom||S|"
    frmCom.Tag2 = "Nº Factura|T|S|||facturassocio|numfactu||S|"
    frmCom.Tag3 = "Fecha Factura|F|N|||facturassocio|fecfactu|dd/mm/yyyy|S|"
    frmCom.Tag4 = "Codigo|N|N|0|999999|facturassocio|codsocio|000000|S|"
    frmCom.Tag5 = "Nombre Socio|T|N|||rsocios|nomsocio||N|"
    
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 7
    frmCom.Maxlen3 = 10
    frmCom.Maxlen4 = 6
    frmCom.Maxlen5 = 35
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "facturassocio"
    frmCom.CampoCP = "numfactu"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Facturas Venta Socio"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|2|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 3000
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaTraspasoAlmacenesPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional EsHistorico As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|NºTraspaso|1400|;S|txtAux(1)|T|Fecha|1405|;S|txtAux(2)|T|Origen|1000|;S|txtAux(3)|T|Almacen Origen|2500|;S|txtAux(4)|T|Destino|995|;S|txtAux(5)|T|Almacen Destino|2500|;"
    
    frmCom.CadenaConsulta = "SELECT scatra.codtrasp, scatra.fechatra, scatra.almaorig, aaa.nomalmac, scatra.almadest, bbb.nomalmac "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM (scatra inner join salmpr aaa on scatra.almaorig = aaa.codalmac) inner join salmpr bbb on scatra.almadest = bbb.codalmac "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    If EsHistorico Then frmCom.CadenaConsulta = Replace(frmCom.CadenaConsulta, "scatra", "schtra")
    
    If Not EsHistorico Then
        frmCom.Tag1 = "NºTraspaso|N|S|0||scatra|codtrasp|0000000|S|"
        frmCom.Tag2 = "Fecha|F|N|||scatra|fechatra|dd/mm/yyyy|N|"
        frmCom.Tag3 = "Origen|N|N|0|999|scatra|almaorig|000|N|"
        frmCom.Tag4 = "Almacen Origen|T|N|||salmpr|nomalmac||N|"
        frmCom.Tag5 = "Destino|N|N|0|999|scatra|almadest|000|N|"
        frmCom.Tag6 = "Almacen Destino|T|N|||salmpr|nomalmac||N|"
    Else
        frmCom.Tag1 = "NºTraspaso|N|S|0||schtra|codtrasp|0000000|S|"
        frmCom.Tag2 = "Fecha|F|N|||schtra|fechatra|dd/mm/yyyy|N|"
        frmCom.Tag3 = "Origen|N|N|0|999|schtra|almaorig|000|N|"
        frmCom.Tag4 = "Almacen Origen|T|N|||salmpr|nomalmac||N|"
        frmCom.Tag5 = "Destino|N|N|0|999|schtra|almadest|000|N|"
        frmCom.Tag6 = "Almacen Destino|T|N|||salmpr|nomalmac||N|"
    End If
    
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 10
    frmCom.Maxlen3 = 3
    frmCom.Maxlen4 = 35
    frmCom.Maxlen5 = 3
    frmCom.Maxlen6 = 35
    
    
    frmCom.pConn = cAgro
    If Not EsHistorico Then
        frmCom.Tabla = "scatra"
    Else
        frmCom.Tabla = "schtra"
    End If
    frmCom.CampoCP = "codtrasp"
    frmCom.TipoCP = "N"
    If Not EsHistorico Then
        frmCom.Caption = "Traspaso de Almacen"
    Else
        frmCom.Caption = "Histórico Traspaso de Almacen"
    End If
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 3000
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaMovimientosAlmacenPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional EsHistorico As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Movimiento|1400|;S|txtAux(1)|T|Fecha|1405|;S|txtAux(2)|T|Origen|1000|;S|txtAux(3)|T|Almacen Origen|2995|;"
    
    frmCom.CadenaConsulta = "SELECT scamov.codmovim, scamov.fecmovim, scamov.codalmac, aaa.nomalmac "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scamov inner join salmpr aaa on scamov.codalmac = aaa.codalmac"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    If EsHistorico Then frmCom.CadenaConsulta = Replace(frmCom.CadenaConsulta, "scamov", "schmov")
    
    If Not EsHistorico Then
        frmCom.Tag1 = "Movimiento|N|S|0||scamov|codmovim|0000000|S|"
        frmCom.Tag2 = "Fecha|F|N|||scamov|fecmovim|dd/mm/yyyy|N|"
        frmCom.Tag3 = "Origen|N|N|0|999|scamov|codalmac|000|N|"
        frmCom.Tag4 = "Almacen Origen|T|N|||salmpr|nomalmac||N|"
    Else
        frmCom.Tag1 = "Movimiento|N|S|0||schmov|codmovim|0000000|S|"
        frmCom.Tag2 = "Fecha|F|N|||schmov|fecmovim|dd/mm/yyyy|N|"
        frmCom.Tag3 = "Origen|N|N|0|999|schmov|codalmac|000|N|"
        frmCom.Tag4 = "Almacen Origen|T|N|||salmpr|nomalmac||N|"
    End If
    
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 10
    frmCom.Maxlen3 = 3
    frmCom.Maxlen4 = 35
    
    
    frmCom.pConn = cAgro
    If Not EsHistorico Then
        frmCom.Tabla = "scamov"
    Else
        frmCom.Tabla = "schmov"
    End If
    frmCom.CampoCP = "codmovim"
    frmCom.TipoCP = "N"
    If Not EsHistorico Then
        frmCom.Caption = "Movimientos de Almacen"
    Else
        frmCom.Caption = "Histórico Movimientos de Almacen"
    End If
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 0
    
    frmCom.Show vbModal
End Sub



Public Sub AyudaMovimientosVariosPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional EsHistorico As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Tipo|2625|;S|txtAux(1)|T|Movimiento|1400|;S|txtAux(2)|T|Fecha|1405|;S|txtAux(3)|T|Socio|1000|;S|txtAux(4)|T|Cliente|1000|;S|txtAux(5)|T|Codigo|800|;S|txtAux(6)|T|Almacen|2995|;"
    
    frmCom.CadenaConsulta = "SELECT case scaser.clisoc when 0 then 'Socio' when 1 then 'Cliente' when 2 then 'Regularizacion Socio' when 3 then 'Regularizacion Cliente' end, scaser.codservi, scaser.fecmovim, scaser.codsocio, scaser.codclien, scaser.codalmac, aaa.nomalmac "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scaser inner join salmpr aaa on scaser.codalmac = aaa.codalmac"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    If EsHistorico Then frmCom.CadenaConsulta = Replace(frmCom.CadenaConsulta, "scaser", "schser")
    
    If Not EsHistorico Then
        frmCom.Tag1 = "Tipo|N|N|||scaser|clisoc||N|"
        frmCom.Tag2 = "Movimiento|N|S|0||scaser|codservi|0000000|S|"
        frmCom.Tag3 = "Fecha|F|N|||scaser|fecmovim|dd/mm/yyyy|N|"
        frmCom.Tag4 = "Socio|N|S|||scaser|codsocio|000000|N|"
        frmCom.Tag5 = "Cliente|N|S|||scaser|codclien|000000|N|"
        frmCom.Tag6 = "Almacen|N|S|||scaser|codalmac|000|N|"
        frmCom.Tag7 = "Nombre Almacen|T|S|||salmpr|nomalmac|||"
    Else
        frmCom.Tag1 = "Tipo|N|N|||schser|clisoc||N|"
        frmCom.Tag2 = "Movimiento|N|S|0||schser|codservi|0000000|S|"
        frmCom.Tag3 = "Fecha|F|N|||schser|fecmovim|dd/mm/yyyy|N|"
        frmCom.Tag4 = "Socio|N|S|||schser|codsocio|000000|N|"
        frmCom.Tag5 = "Cliente|N|S|||schser|codclien|000000|N|"
        frmCom.Tag6 = "Almacen|N|S|||schser|codalmac|000|N|"
        frmCom.Tag7 = "Nombre Almacen|T|S|||salmpr|nomalmac|||"
    End If
    
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 10
    frmCom.Maxlen3 = 3
    frmCom.Maxlen4 = 35
    
    
    frmCom.pConn = cAgro
    If Not EsHistorico Then
        frmCom.Tabla = "scaser"
    Else
        frmCom.Tabla = "schser"
    End If
    frmCom.CampoCP = "codservi"
    frmCom.TipoCP = "N"
    If Not EsHistorico Then
        frmCom.Caption = "Movimientos de Servicios Varios"
    Else
        frmCom.Caption = "Histórico Movimientos de Servicios Varios"
    End If
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 4225
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaHistoricoInventarioPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Artículo|2205|;S|txtAux(1)|T|Nombre|4795|;"
    
    frmCom.CadenaConsulta = "SELECT distinct shinve.codartic, sartic.nomartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM shinve inner join sartic on shinve.codartic = sartic.codartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Artículo|T|N|||shinve|codartic||S|"
    frmCom.Tag2 = "Denominacion|T|N|||sartic|nomartic||N|"
    
    frmCom.Maxlen1 = 16
    frmCom.Maxlen2 = 35
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "shinve"
    frmCom.CampoCP = "codartic"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Histórico Inventario"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 0
    
    frmCom.Show vbModal
End Sub








Private Sub Redimensiona(frmBas As frmBasico2, Cant As Integer)
    frmBas.Width = frmBas.Width + Cant
    frmBas.DataGrid1.Width = frmBas.DataGrid1.Width + Cant
    frmBas.cmdAceptar.Left = frmBas.cmdAceptar.Left + Cant
    frmBas.cmdCancelar.Left = frmBas.cmdCancelar.Left + Cant
    frmBas.cmdRegresar.Left = frmBas.cmdRegresar.Left + Cant

End Sub


