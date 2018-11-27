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






Private Sub Redimensiona(frmBas As frmBasico2, Cant As Integer)
    frmBas.Width = frmBas.Width + Cant
    frmBas.DataGrid1.Width = frmBas.DataGrid1.Width + Cant
    frmBas.cmdAceptar.Left = frmBas.cmdAceptar.Left + Cant
    frmBas.cmdCancelar.Left = frmBas.cmdCancelar.Left + Cant
    frmBas.cmdRegresar.Left = frmBas.cmdRegresar.Left + Cant

End Sub

