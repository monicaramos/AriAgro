Attribute VB_Name = "ModBasico"
Option Explicit


Public Sub AyudaFamilias(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|905|;S|txtAux(1)|T|Descripci�n|4095|;S|txtAux(2)|T|Cta.Ventas|1500|;S|txtAux(3)|T|Cta.Compras|1500|;"
    frmCom.CadenaConsulta = "SELECT sfamia.codfamia, sfamia.nomfamia, sfamia.ctaventa, sfamia.ctacompr "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sfamia "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "C�digo|N|N|0|999|sfamia|codfamia|000|S|"
    frmCom.Tag2 = "Descripci�n|T|N|||sfamia|nomfamia|||"
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
    
    frmCom.Tag1 = "N� Palet|N|S|||palets|numpalet|0000000|S|"
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
    frmCom.Tag2 = "N� Factura|T|N|||tcafpc|numfactu||S|"
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
    
    frmCom.Tag1 = "N�Albar�n|N|S|||scaalb|numalbar|0000000|S|"
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

    frmCom.CadenaTots = "S|txtAux(0)|T|Cliente|1100|;S|txtAux(1)|T|Nombre|4795|;S|txtAux(2)|T|NIF|1905|;S|txtAux(3)|T|Telefono|1900|;S|txtAux(3)|T|M�vil|1900|;"
    
    frmCom.CadenaConsulta = "SELECT clientes.codclien, clientes.nomclien, clientes.cifclien, clientes.telclie1, clientes.movclie1 "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM clientes  "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    
    frmCom.Tag1 = "Codigo|N|S|0|999999|clientes|codclien|000000|S|"
    frmCom.Tag2 = "Nombre|T|S|||clientes|nomclien|||"
    frmCom.Tag3 = "NIF|T|S|||clientes|cifclien|||"
    frmCom.Tag4 = "Tel�fono|T|S|||clientes|telclie1|||"
    frmCom.Tag5 = "M�vil|T|S|||clientes|movclie1|||"
    
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
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|1905|;S|txtAux(1)|T|Nombre|4195|;S|txtAux(2)|T|Codigo EAN|1900|;"
    frmCom.CadenaConsulta = "SELECT sartic.codartic, sartic.nomartic, sartic.codigoea"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "C�digo|T|N|||sartic|codartic|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||sartic|nomartic|||"
    frmCom.Tag3 = "C�digo EAN|T|S|||sartic|codigoea|||"
    
    frmCom.Maxlen1 = 16
    frmCom.Maxlen2 = 40
    frmCom.Maxlen3 = 13
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "sartic"
    frmCom.CampoCP = "codartic"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Art�culos"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 1000
    
    frmCom.Show vbModal
End Sub




Public Sub AyudaForfaits(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|1905|;S|txtAux(1)|T|Descripci�n|5095|;S|txtAux(2)|T|Variedad|1000|;S|txtAux(3)|T|Nombre Variedad|2500|;"
    frmCom.CadenaConsulta = "SELECT forfaits.codforfait, forfaits.nomconfe, forfaits.codvarie, variedades.nomvarie "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM forfaits left join variedades on forfaits.codvarie = variedades.codvarie "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "C�digo forfait|T|N|||forfaits|codforfait||S|"
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
    frmCom.Tag4 = "Tel�fono|T|S|||proveedor|telprov1|||"
    frmCom.Tag5 = "M�vil|T|S|||proveedor|faxprov1|||"
    
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
    frmCom.CadenaTots = "S|txtAux(0)|T|Pedido|1105|;S|txtAux(1)|T|Fecha|1400|;S|txtAux(2)|T|C�digo|1000|;S|txtAux(3)|T|Proveedor|4595|;"
    
    frmCom.CadenaConsulta = "SELECT scappr.numpedpr, scappr.fecpedpr, scappr.codprove, scappr.nomprove "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scappr "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "N� Pedido|N|S|0||scappr|numpedpr|0000000|S|"
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
    frmCom.CadenaTots = "S|txtAux(0)|T|Albaran|1105|;S|txtAux(1)|T|Fecha|1400|;S|txtAux(2)|T|C�digo|1000|;S|txtAux(3)|T|Proveedor|4595|;"
    
    frmCom.CadenaConsulta = "SELECT scaalp.numalbar, scaalp.fechaalb, scaalp.codprove, scaalp.nomprove "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scaalp "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "N� Albaran|N|S|0||scaalp|numalbar|0000000|S|"
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
    frmCom.CadenaTots = "S|txtAux(0)|T|Factura|1405|;S|txtAux(1)|T|Fecha|1400|;S|txtAux(2)|T|F.Recepci�n|1400|;S|txtAux(3)|T|C�digo|1000|;S|txtAux(4)|T|Proveedor|4795|;"
    
    frmCom.CadenaConsulta = "SELECT scafpc.numfactu, scafpc.fecfactu, scafpc.fecrecep, scafpc.codprove, scafpc.nomprove "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scafpc "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "N� Factura|T|S|||scafpc|numfactu||S|"
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
    frmCom.CadenaTots = "S|txtAux(0)|T|Tipo|1400|;S|txtAux(1)|T|Factura|1405|;S|txtAux(2)|T|Fecha|1400|;S|txtAux(3)|T|C�digo|1000|;S|txtAux(4)|T|Socio|4795|;"
    
    frmCom.CadenaConsulta = "SELECT facturassocio.codtipom, facturassocio.numfactu, facturassocio.fecfactu, facturassocio.codsocio, rsocios.nomsocio "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM facturassocio inner join rsocios on facturassocio.codsocio = rsocios.codsocio "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Tipo Movimiento|T|N|||facturassocio|codtipom||S|"
    frmCom.Tag2 = "N� Factura|T|S|||facturassocio|numfactu||S|"
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
    frmCom.CadenaTots = "S|txtAux(0)|T|N�Traspaso|1400|;S|txtAux(1)|T|Fecha|1405|;S|txtAux(2)|T|Origen|1000|;S|txtAux(3)|T|Almacen Origen|2500|;S|txtAux(4)|T|Destino|995|;S|txtAux(5)|T|Almacen Destino|2500|;"
    
    frmCom.CadenaConsulta = "SELECT scatra.codtrasp, scatra.fechatra, scatra.almaorig, aaa.nomalmac, scatra.almadest, bbb.nomalmac "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM (scatra inner join salmpr aaa on scatra.almaorig = aaa.codalmac) inner join salmpr bbb on scatra.almadest = bbb.codalmac "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    If EsHistorico Then frmCom.CadenaConsulta = Replace(frmCom.CadenaConsulta, "scatra", "schtra")
    
    If Not EsHistorico Then
        frmCom.Tag1 = "N�Traspaso|N|S|0||scatra|codtrasp|0000000|S|"
        frmCom.Tag2 = "Fecha|F|N|||scatra|fechatra|dd/mm/yyyy|N|"
        frmCom.Tag3 = "Origen|N|N|0|999|scatra|almaorig|000|N|"
        frmCom.Tag4 = "Almacen Origen|T|N|||salmpr|nomalmac||N|"
        frmCom.Tag5 = "Destino|N|N|0|999|scatra|almadest|000|N|"
        frmCom.Tag6 = "Almacen Destino|T|N|||salmpr|nomalmac||N|"
    Else
        frmCom.Tag1 = "N�Traspaso|N|S|0||schtra|codtrasp|0000000|S|"
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
        frmCom.Caption = "Hist�rico Traspaso de Almacen"
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
        frmCom.Caption = "Hist�rico Movimientos de Almacen"
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
        frmCom.Caption = "Hist�rico Movimientos de Servicios Varios"
    End If
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 4225
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaHistoricoInventarioPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Art�culo|2205|;S|txtAux(1)|T|Nombre|4795|;"
    
    frmCom.CadenaConsulta = "SELECT distinct shinve.codartic, sartic.nomartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM shinve inner join sartic on shinve.codartic = sartic.codartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Art�culo|T|N|||shinve|codartic||S|"
    frmCom.Tag2 = "Denominacion|T|N|||sartic|nomartic||N|"
    
    frmCom.Maxlen1 = 16
    frmCom.Maxlen2 = 35
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "shinve"
    frmCom.CampoCP = "codartic"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Hist�rico Inventario"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 0
    
    frmCom.Show vbModal
End Sub



Public Sub AyudaVariedad(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|1405|;S|txtAux(1)|T|Descripci�n|3000|;S|txtAux(2)|T|Producto|2595|;"
    frmBas.CadenaConsulta = "SELECT variedades.codvarie, variedades.nomvarie, productos.nomprodu "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM variedades inner join productos on variedades.codprodu = productos.codprodu "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "C�digo |N|N|||variedades|codvarie|000000|S|"
    frmBas.Tag2 = "Descripci�n|T|N|||variedades|nomvarie|||"
    frmBas.Tag3 = "Producto|T|N|||variedades|nomprodu|||"
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 30
    frmBas.Maxlen3 = 100
    
    frmBas.pConn = cAgro
    
    frmBas.Tabla = "variedades"
    frmBas.CampoCP = "codvarie"
    frmBas.TipoCP = "N"
    frmBas.Caption = "Variedades"
    frmBas.DeConsulta = True
    
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    '[Monica]17/04/2018: a�adimos en este caso los botones de busqueda
    frmBas.DataGrid1.Height = 7420
    frmBas.DataGrid1.Top = 870
    frmBas.FrameBotonGnral.visible = True
    frmBas.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    frmBas.Show vbModal

    
End Sub


Public Sub AyudaFacturasComprasPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Factura|1400|;S|txtAux(1)|T|Fecha|1400|;S|txtAux(2)|T|Codigo|1000|;S|txtAux(3)|T|Proveedor|4500|;S|txtAux(4)|T|F.Recepcion|1400|;"
    
    frmCom.CadenaConsulta = "SELECT facturascom.numfactu, facturascom.fecfactu, facturascom.codprove, proveedor.nomprove, facturascom.fecrecep"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM facturascom inner join proveedor on facturascom.codprove = proveedor.codprove "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Proveedor|N|N|0|999999|facturascom|codprove|000000|S|"
    frmCom.Tag2 = "N� Factura|T|S|||facturascom|numfactu||S|"
    frmCom.Tag3 = "Proveedor|N|N|0|999999|facturascom|codprove|000000|S|"
    frmCom.Tag4 = "Nombre Proveedor|T|N|||proveedor|nomprove|||"
    frmCom.Tag5 = "Fecha Recepci�n|F|N|||facturascom|fecrecep|dd/mm/yyyy||"
    
    frmCom.Maxlen1 = 10
    frmCom.Maxlen2 = 10
    frmCom.Maxlen3 = 6
    frmCom.Maxlen4 = 35
    frmCom.Maxlen5 = 10
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "facturascom"
    frmCom.CampoCP = "numfactu"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Facturas Compras"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|2|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 2800
    
    frmCom.Show vbModal
End Sub

 
Public Sub AyudaAlmMovArtPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Art�culo|2205|;S|txtAux(1)|T|Nombre|4795|;"
    
    frmCom.CadenaConsulta = "SELECT distinct smoval.codartic, sartic.nomartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM smoval inner join sartic on smoval.codartic = sartic.codartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Art�culo|T|N|||smoval|codartic||S|"
    frmCom.Tag2 = "Denominacion|T|N|||sartic|nomartic||N|"
    
    frmCom.Maxlen1 = 16
    frmCom.Maxlen2 = 35
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "smoval"
    frmCom.CampoCP = "codartic"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Movimientos Art�culos"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 0
    
    frmCom.Show vbModal
End Sub

     
Public Sub AyudaAneExpedPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Expediente|2705|;S|txtAux(1)|T|Linea|1295|;S|txtAux(2)|T|Campa�a|1500|;S|txtAux(3)|T|Periodo|1500|;"
    
    frmCom.CadenaConsulta = "SELECT anecoop.expediente_id, anecoop.linea_expediente, anecoop.codigo_campanya, anecoop.periodo "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM anecoop left join anecoop_pago on anecoop.expediente_id = anecoop_pago.expediente_id "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Expediente|T|S|||anecoop|expediente_id||S|"
    frmCom.Tag2 = "Linea|T|S|||anecoop|linea_expediente||S|"
    frmCom.Tag3 = "Codigo Campa�a|T|S|||anecoop|codigo_campanya||S|"
    frmCom.Tag4 = "Periodo|T|S|||anecoop|periodo|||"
    
    frmCom.Maxlen1 = 16
    frmCom.Maxlen2 = 10
    frmCom.Maxlen3 = 5
    frmCom.Maxlen3 = 5
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "anecoop left join anecoop_pago on anecoop.expediente_id = anecoop_pago.expediente_id"
    frmCom.CampoCP = "anecoop.expediente_id"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Expedientes de ANECOOP"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 0
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaDireccionesPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|905|;S|txtAux(1)|T|Tipo|1095|;S|txtAux(2)|T|Nombre Direccion|5000|;"
    frmCom.CadenaConsulta = "SELECT sdirpr.coddirec, if(tipodire=0,""Albaran"",""Factura"") as tipodire, sdirpr.nomdirec "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sdirpr "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Cod. Direcci�n|N|N|0|999|sdirpr|coddirec|000|S|"
    frmCom.Tag2 = "Tipo Direcci�n|N|N|||sdirpr|tipodire||N|"
    frmCom.Tag3 = "Nombre Direcci�n|T|N|||sdirpr|nomdirec||N|"
    
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Maxlen3 = 30
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "sdirpr"
    frmCom.CampoCP = "coddirec"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Direcciones de compras"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 0
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaProveVPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|NIF|2000|;S|txtAux(1)|T|Nombre Proveedor|5000|;"
    frmCom.CadenaConsulta = "SELECT sprvar.nifprove, sprvar.nomprove "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sprvar "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "N.I.F.|T|N|||sprvar|nifprove||S|"
    frmCom.Tag2 = "Nombre Proveedor Varios|T|N|||sprvar|nomprove||N|"
    
    frmCom.Maxlen1 = 15
    frmCom.Maxlen2 = 30
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "sprvar"
    frmCom.CampoCP = "nifprove"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Proveedores Varios"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 0
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaPrecProvPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|1205|;S|txtAux(1)|T|Proveedor|3295|;S|txtAux(2)|T|Codigo|1500|;S|txtAux(3)|T|Art�culo|3000|;"
    
    frmCom.CadenaConsulta = "SELECT slispr.codprove, proveedor.nomprove, slispr.codartic, sartic.nomartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM (slispr left join proveedor on slispr.codprove = proveedor.codprove) left join sartic on slispr.codartic = sartic.codartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Cod. Proveedor|N|N|0|999999|slispr|codprove|000000|S|"
    frmCom.Tag2 = "Proveedor|T|N|||proveedor|nomprove|||"
    frmCom.Tag3 = "Cod. Art�culo|T|N|||slispr|codartic||S|"
    frmCom.Tag4 = "Art�culo|T|S|||sartic|nomartic|||"
    
    frmCom.Maxlen1 = 6
    frmCom.Maxlen2 = 35
    frmCom.Maxlen3 = 16
    frmCom.Maxlen3 = 35
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "(slispr left join proveedor on slispr.codprove = proveedor.codprove) left join sartic on slispr.codartic = sartic.codartic "
    frmCom.CampoCP = "slispr.codprove"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Precios Proveedor"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|2|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 2000
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaNomCostePrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Codigo|1000|;S|txtAux(1)|T|Nombre|5000|;"
    frmCom.CadenaConsulta = "SELECT nombcoste.codcoste, nombcoste.denominacion "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM nombcoste "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "C�digo|N|N|||nombcoste|codcoste|00|S|"
    frmCom.Tag2 = "Nombre|T|N|||nombcoste|denominacion|||"
    
    frmCom.Maxlen1 = 15
    frmCom.Maxlen2 = 30
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "nombcoste"
    frmCom.CampoCP = "codcoste"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Denominaci�n de Costes"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, -1000
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaPaletConfPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Codigo|1000|;S|txtAux(1)|T|Descripcion|3800|;S|txtAux(2)|T|Peso|1200|;"
    frmCom.CadenaConsulta = "SELECT confpale.codpalet, confpale.nompalet, confpale.pesopale "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM confpale "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "C�digo|N|N|0|999|confpale|codpalet|000|S|"
    frmCom.Tag2 = "Descripci�n|T|N|||confpale|nompalet|||"
    frmCom.Tag3 = "Peso Palet|N|N|||confpale|pesopale|#0.00||"
    
    
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 30
    frmCom.Maxlen3 = 10
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "confpale"
    frmCom.CampoCP = "codpalet"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Paletizacion"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, -1000
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaAgenciasPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Codigo|1000|;S|txtAux(1)|T|Descripcion|4200|;S|txtAux(2)|T|Tipo|1800|;"
    frmCom.CadenaConsulta = "SELECT agencias.codtrans, agencias.nomtrans, if(agencias.tipo=0,'Transportista','Comisionista') "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM agencias "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "C�digo Agencia|N|N|0|999|agencias|codtrans|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||agencias|nomtrans|||"
    frmCom.Tag3 = "Tipo|T|N|||agencias|tipo|||"
    
    
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 30
    frmCom.Maxlen3 = 10
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "agencias"
    frmCom.CampoCP = "codtrans"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Agencias de transporte / Comisionistas"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 0
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaCCostesDiarisoPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Fecha|1400|;S|txtAux(1)|T|Coste|1000|;S|txtAux(2)|T|Descripcion|3800|;S|txtAux(3)|T|Observaciones|3800|;"
    frmCom.CadenaConsulta = "SELECT cccabdia.fecha, cccabdia.codcoste, ccconcostes.nomcoste, cccabdia.observac "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM cccabdia "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Fecha|F|N|||cccabdia|fecha|dd/mm/yyyy|S|"
    frmCom.Tag2 = "C�digo Coste|N|S|||cccabdia|codcoste|000000|S|"
    frmCom.Tag3 = "Nombre Coste|N|S|||ccconcostes|nomcoste|||"
    frmCom.Tag4 = "Observaciones|T|S|||cccabdia|observac|||"
    
    
    frmCom.Maxlen1 = 10
    frmCom.Maxlen2 = 6
    frmCom.Maxlen3 = 30
    frmCom.Maxlen4 = 30
    
    frmCom.pConn = cAgro
    
    frmCom.Tabla = "cccabdia"
    frmCom.CampoCP = "fecha"
    frmCom.TipoCP = "F"
    frmCom.Caption = "Entrada Costes Diarios"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 3000
    
    frmCom.Show vbModal
    
End Sub

Private Sub Redimensiona(frmBas As frmBasico2, Cant As Integer)
    frmBas.Width = frmBas.Width + Cant
    frmBas.DataGrid1.Width = frmBas.DataGrid1.Width + Cant
    frmBas.cmdAceptar.Left = frmBas.cmdAceptar.Left + Cant
    frmBas.cmdCancelar.Left = frmBas.cmdCancelar.Left + Cant
    frmBas.cmdRegresar.Left = frmBas.cmdRegresar.Left + Cant

End Sub


