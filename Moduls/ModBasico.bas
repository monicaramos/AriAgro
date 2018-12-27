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
    
    frmCom.tabla = "sfamia"
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
    
    frmCom.tabla = "palets inner join confpale ON palets.codpalet=confpale.codpalet "
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
    
    frmCom.tabla = "albaran inner join clientes ON albaran.codclien=clientes.codclien "
    frmCom.CampoCP = "numalbar"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Albaranes"
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
    
    frmCom.tabla = "pedidos inner join clientes ON pedidos.codclien=clientes.codclien "
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
    
    frmCom.tabla = "facturas inner join clientes ON facturas.codclien=clientes.codclien "
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
    
    frmCom.tabla = "tcafpc inner join agencias ON tcafpc.codtrans=agencias.codtrans "
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

    frmCom.CadenaTots = "S|txtAux(0)|T|Albaran|1505|;S|txtAux(1)|T|Cliente|1100|;S|txtAux(3)|T|Nombre|4795|;S|txtAux(2)|T|Fecha|1400|;"
    
    frmCom.CadenaConsulta = "SELECT scaalb.numalbar, scaalb.codclien, clientes.nomclien, scaalb.fecalbar "
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
    
    frmCom.tabla = "scaalb inner join clientes ON scaalb.codclien=clientes.codclien "
    frmCom.CampoCP = "numalbar"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Albaranes de Envases"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 1800
    
    frmCom.Show vbModal
End Sub


Private Sub Redimensiona(frmBas As frmBasico2, Cant As Integer)
    frmBas.Width = frmBas.Width + Cant
    frmBas.DataGrid1.Width = frmBas.DataGrid1.Width + Cant
    frmBas.cmdAceptar.Left = frmBas.cmdAceptar.Left + Cant
    frmBas.cmdCancelar.Left = frmBas.cmdCancelar.Left + Cant
    frmBas.cmdRegresar.Left = frmBas.cmdRegresar.Left + Cant

End Sub

