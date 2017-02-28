Attribute VB_Name = "modMenuClick"
Option Explicit

Dim DeTransporte As Boolean
Dim DeServicios As Boolean
Dim frmBas As frmBasico

Private Sub Construc(nom As String)
    MsgBox nom & ": en construcció..."
End Sub

' ******* DATOS BASICOS *********

Public Sub SubmnP_Generales_Click(Index As Integer)

    Select Case Index
        Case 1: frmConfParamGral.Show vbModal
                PonerDatosPpal
        Case 2: frmConfParamAplic.Show vbModal
        Case 3: conn.Close
                If AbrirConexionUsuarios Then
                    frmConfTipoMov.Show vbModal
                    CerrarConexionUsuarios
                End If
                If AbrirConexion() = False Then
                    MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
                    End
                End If
        Case 4: frmConfParamRpt.Show vbModal
        Case 6: frmMantenusu.Show vbModal ' mantenimiento de usuarios
        Case 8: frmBD.Show vbModal 'creacion de una nueva base de datos
        Case 9: 'traspaso a la tabla de chivato
'                If Dir(App.path & "\ExportChivato.exe", vbArchive) <> "" And _
'                   Dir(App.path & "\ConfigChivato.ini", vbArchive) <> "" Then
'                    Shell App.path & "\ExportChivato.exe", vbNormalFocus
'                Else
'                    MsgBox "No tiene los ficheros necesarios. LLame a Ariadna.", vbExclamation
'                End If
'[Monica]13/12/2011:Embebemos el proceso de carga de chivato en el propio programa
               frmChivato.Show vbModal

        Case 11: End
    End Select
End Sub

Public Sub submnM_Generales_click(Index As Integer)
    Select Case Index
        Case 1: frmManAlmProp.Show vbModal 'almacenes propios
        Case 2: frmManTipUnid.Show vbModal 'tipos de unidad
        Case 3: frmManTipArtic.Show vbModal 'tipos de articulos
        Case 4: frmManFamilias.Show vbModal 'familias
        Case 5: frmManArtic.Show vbModal 'articulos
    End Select
End Sub


' *******  COMERCIAL  *********

Public Sub SubmnC_ComercialG_Admon_Click(Index As Integer)
    Select Case Index
        Case 1: frmManPaises.Show vbModal ' paises
        Case 2: frmManTipMerc.Show vbModal 'tipos de mercados
        Case 3: frmManCadenas.Show vbModal ' cadenas de supermercados
        Case 4: frmManAgencias.Show vbModal ' agencias
        Case 5: frmManFpago.Show vbModal 'formas de pago
'        Case 6: frmManClien.Show vbModal ' clientes
        Case 6: frmClientes.Show vbModal ' clientes
        Case 7: frmManBanco.Show vbModal  ' bancos propios
        Case 8: frmFacCartasOferta.Show vbModal ' tipos de cartas
        Case 9: frmManInciden.Show vbModal ' incidencias
        Case 10: frmTipIVAConta.Show vbModal  ' tipos de iva

' el mantenimiento de tarifas estará en la parte de recoleccion.
'               frmManTarifas.Show vbModal ' tarifas

    End Select
End Sub

Public Sub SubmnC_ComercialG_Produc_Click(Index As Integer)
    Select Case Index
        Case 1: frmManGrupos.Show vbModal  'grupos de productos
        Case 2: frmManProductos.Show vbModal 'productos
        Case 3: frmManClases.Show vbModal 'clases
        Case 4: frmManVariedad.Show vbModal 'variedades
        Case 5: frmManCalibres.Show vbModal 'calibres
        Case 6: frmManMarcas.Show vbModal 'marcas
        Case 7: frmManCodEAN.Show vbModal 'codigos ean
    End Select
End Sub


Public Sub SubmnC_ComercialG_Confec_Click(Index As Integer)
    Select Case Index
        Case 1: frmManTipEnv.Show vbModal ' tipos de Envases
        Case 2: frmManCapEnv.Show vbModal ' capacidad de envases
        Case 3: frmManMedEnv.Show vbModal ' medidas de envases
        Case 4: frmManFConf.Show vbModal  ' formas de confeccion
        Case 5: frmManPresConf.Show vbModal  ' presentaciones de confeccion
        Case 6: frmManPaletConf.Show vbModal  ' paletizacion de confeccion
        Case 7: frmManNomCoste.Show vbModal  ' nombre de coste
        Case 8: frmManForfaits.Show vbModal  ' forfaits --> confecciones
    End Select
End Sub

' *******  GESTION DE MATERIALES  *********
' *******  DATOS GENERALES  *********

Public Sub SubmnC_GM_Gral_Click(Index As Integer)
    Select Case Index
        Case 1: frmManAlmProp.Show vbModal 'almacenes propios
        Case 2: frmManTipUnid.Show vbModal 'tipos de unidad
        Case 3: frmManTipArtic.Show vbModal 'tipos de articulos
        Case 4: frmManFamilias.Show vbModal 'familias
        Case 5: frmManArtic.Show vbModal 'articulos
        Case 7: frmManProve.Show vbModal ' proveedores
        Case 8: frmComProveV.Show vbModal  'proveedores varios
        Case 9: frmComDirecciones.Show vbModal  'direcciones
        Case 11: frmComPreciosProv.Show vbModal ' precios proveedor
    End Select
End Sub

' *******  GESTION DE MATERIALES  *********
' *******  MOVIMIENTOS DE ALMACEN  *********

Public Sub SubmnC_GM_MovAlm_Click(Index As Integer)
    Select Case Index
        Case 1:  'traspaso de almacenes
                frmAlmTraspaso.EsHistorico = False
                frmAlmTraspaso.hcoCodMovim = -1
                frmAlmTraspaso.Show vbModal
        
        Case 2:  'hco traspaso de almacenes
                frmAlmTraspaso.EsHistorico = True
                frmAlmTraspaso.hcoCodMovim = -1
                frmAlmTraspaso.Show vbModal
        
        Case 3:  'movimientos almacen
                frmAlmMovimientos.EsHistorico = False
                frmAlmMovimientos.hcoCodMovim = -1 'No carga el form al abrir
                frmAlmMovimientos.Show vbModal
        
        Case 4:  'hco movimientos de almacen
                frmAlmMovimientos.EsHistorico = True
                frmAlmMovimientos.hcoCodMovim = -1
                frmAlmMovimientos.Show vbModal
                
        Case 5:  ' movimientos de servicios varios
                frmAlmMovimientosVar.EsHistorico = False
                frmAlmMovimientosVar.hcoCodMovim = -1 'No carga el form al abrir
                frmAlmMovimientosVar.Show vbModal
        
        Case 6:  ' hco de movimientos de servicios varios
                frmAlmMovimientosVar.EsHistorico = True
                frmAlmMovimientosVar.hcoCodMovim = -1
                frmAlmMovimientosVar.Show vbModal
        
                
        Case 8: 'Mantenimiento de Envases Retornables
                frmAlmEnvRet.Show vbModal
                
        Case 9: ' Grabación de ficheros cheps
                frmAlmGrabChep.Show vbModal
    End Select
End Sub

' *******  GESTION DE MATERIALES  *********
' *******  CONSULTAS  *********

Public Sub SubmnC_GM_Consul_Click(Index As Integer)
    Select Case Index
        Case 1:  'movimientos de articulos
                frmAlmMovimArticulos.Show vbModal
        Case 2:  'listado de movimientos
                DeServicios = False
                AbrirListado2 (9)
        Case 3:  'listado de valoracion de stocks
                AbrirListado2 (17)
        Case 4:  'informe de stocks max_min
                AbrirListado2 (18)
        Case 5:  'informe de stocks a una fecha
                AbrirListado2 (19)
        Case 6:  'informe de movimiento de articulos por familia
                frmListMovArtFam.Show vbModal
        Case 7: ' Informe de movimientos por socio/cliente
                DeServicios = True
                AbrirListado2 (9)
    End Select
End Sub

' *******  GESTION DE MATERIALES  *********
' *******  INFORMES VARIOS  *********

Public Sub SubmnC_GM_InfVa_Click(Index As Integer)
    Select Case Index
        Case 1:  'informe de proveedores
                AbrirListado (11)
        Case 2:  'etiquetas de proveedores
                AbrirListadoOfer (305)
        Case 3:  'cartas a proveedores
                AbrirListadoOfer (306)
    End Select
End Sub



' *******  GESTION DE MATERIALES  *********
' *******  INVENTARIO  *********

Public Sub SubmnC_GM_Inven_Click(Index As Integer)
    Select Case Index
        Case 1:  'toma de inventario
                AbrirListado2 (12)
        Case 2:  'entrada de existencia real
                frmAlmInventario.Show vbModal
        Case 3:  'listado de diferencias
                AbrirListado2 (13)
        Case 4:  'actualizar diferencias
                AbrirListado2 (14)
        Case 5:  'valoracion de stocks inventariados
                AbrirListado2 (16)
        Case 6:  'historico inventario
                frmAlmHcoInven.Show vbModal
                
        Case 8:  'Recálculo del precio medio ponderado
                AbrirListado2 (120)
                
    End Select
End Sub

' *******  GESTION DE MATERIALES  *********
' *******  PEDIDOS PROVEEDOR  *********

Public Sub SubmnC_GM_PedCom_Click(Index As Integer)
    Select Case Index
        Case 1: 'pedidos de proveedor
                frmComEntPedidos.EsHistorico = False
                frmComEntPedidos.Show vbModal
        Case 2: 'historico de pedidos anulados
                frmComEntPedidos.EsHistorico = True
                frmComEntPedidos.Show vbModal
        Case 3: 'listado de material pendiente de recibir
                AbrirListadoOfer (307) '307: List. Materia pte recibir
    End Select
End Sub


' *******  GESTION DE MATERIALES  *********
' *******  ALBARANES PROVEEDOR  *********

Public Sub SubmnC_GM_AlbCom_Click(Index As Integer)
    Select Case Index
        Case 1: 'albaranes proveedor
                frmComEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
                frmComEntAlbaranes.EsHistorico = False
                frmComEntAlbaranes.Show vbModal
        Case 2: 'Historico albaranes de compras a proveedores
                frmComEntAlbaranes.EsHistorico = True
                frmComEntAlbaranes.Show vbModal
        Case 3: 'Listado de Albaranes pendientes de Factura
                AbrirListadoOfer (308) '308: List. Albaranes pte facturar
        Case 5: 'Recepción de facturas de proveedor
                frmComFacturar.Show vbModal
        Case 6: 'historico de facturas
                frmComHcoFacturas.hcoCodMovim = ""
                frmComHcoFacturas.Show vbModal
        Case 8: 'Contabilizar Facturas
                AbrirListado2 (224) 'Para pedir datos
        
    End Select
End Sub


' *******  GESTION DE MATERIALES  *********
' *******  ESTADISTICAS PROVEEDOR  *********

Public Sub SubmnC_GM_EstCom_Click(Index As Integer)
    Select Case Index
        Case 1: 'Listado de compras por proveedor
                AbrirListadoOfer (310)
        Case 2: 'Listado de compras por Familia
                AbrirListadoOfer (311)
    End Select
End Sub


' *******  GESTION DE VENTAS  *********
' *******  CARGA AUTOMATICA PALETS *********

Public Sub SubmnC_GV_Palets_Click(Index As Integer)
    Select Case Index
        Case 1: frmManMovimTRZ.Show vbModal
        'AbrirListado (1) 'Creacion automatica de palets
    End Select
End Sub


' *******  GESTION DE VENTAS  *********
' *******  PALETS Y PEDIDOS  *********



Public Sub SubmnC_GV_PalPed_Click(Index As Integer)
    Select Case Index
        Case 1: frmVtasPalets.Show vbModal 'Mantenimiento de palets
        Case 2: frmListPaletsConf.Show vbModal   'frmVtasPaletsConf.Show vbModal 'Listado de palets confeccionados
        Case 4: frmVtasPedidos.Show vbModal 'Mantenimiento de pedidos
        Case 5: frmVtasPlanDiario.Show vbModal 'planning de Confeccion
        Case 6: frmVtasOrdenConfec.Show vbModal  'orden de Confeccion
        Case 8: frmVtasAlbaranes.Show vbModal 'Mantenimiento de albaranes
        Case 9: frmVtasRevalorarAlb.Show vbModal 'Revalorar costes de albaranes
        Case 10: frmVtasManCostReal.Show vbModal 'frmVtasRevalorarCostes.Show vbModal 'Revalorar Costes Reales
        Case 11: frmVtasAlbPdtesFact.Show vbModal 'albaranes pendientes de facturar
        Case 13: frmImpAridoc.Tipo = 0 ' Integracion de aridoc: Albaranes de venta
                 frmImpAridoc.Caption = "Importar Albaranes de Venta a Aridoc"
                 frmImpAridoc.Label4(16).Caption = "Fecha Albarán"
                 frmImpAridoc.Show vbModal 'vbModalConstruc("Integracion aridoc")
        Case 15: Select Case vParamAplic.Cooperativa
                    Case 1 ' Valsur
                        frmTrazabilidad.OpcionListado = 0
                        frmTrazabilidad.Show vbModal 'traspaso de trazabilidad
                    Case 5 ' Castelduc
                        frmTrazabilidad.OpcionListado = 1
                        frmTrazabilidad.Show vbModal 'traspaso de trazabilidad
                 End Select
    End Select
End Sub

' *******  GESTION DE VENTAS  *********
' *******  ALBARANES ENVASES   ********

Public Sub SubmnC_GV_AlbEnv_Click()
'    Construc ("Albaranes Envases") 'frmVtasAlbEnvases.Show vbModal ' Informes de Gestion de Ventas Gráficos
    frmVtasAlbEnvases.Show vbModal ' Informes de Gestion de Ventas Gráficos
End Sub


' *******         GESTION DE VENTAS       *********
' *******  FACTURACION ALBARANES ENVASES   ********

Public Sub SubmnC_GV_FactAlbEnv_Click()
'    Construc ("Facturacion Albaranes Envases") 'frmVtasAlbEnvases.Show vbModal ' Informes de Gestion de Ventas Gráficos
    frmVtasFactAlbEnv.OpcionListado = 52 'facturacion de albaranes
    frmVtasFactAlbEnv.Show vbModal
End Sub


' *******  GESTION DE VENTAS  *********
' *******    INFORMES         ********

Public Sub SubmnC_GV_Informes_Click()
    frmVtasInformes.Show vbModal ' Informes de Gestion de Ventas
End Sub

' *******  GESTION DE VENTAS  *********
' *******    INTRASTAT         ********

Public Sub SubmnC_GV_Intrastat_Click()
    frmVtasIntrastat.Show vbModal ' Informe de intrastat
End Sub



' *******  GESTION DE VENTAS  *********
' *******  RDTO POR VARIEDAD  ********

Public Sub SubmnC_GV_Rdto_Click()
    frmVtasRdtoVar.Show vbModal ' Informe de Rendimiento por Variedad
End Sub

' *******  GESTION DE VENTAS  *********
' *******  INFORMES GRAFICOS   ********

Public Sub SubmnC_GV_InfGra_Click()
    frmVtasInfGra.Show vbModal ' Informes de Gestion de Ventas Gráficos
End Sub

' *******  GESTION DE VENTAS  *********
' *******    DIFERENCIAS       ********

Public Sub SubmnC_GV_Diferencias_Click()
    frmVtasListDif.Show vbModal ' Informe de Diferencias entre kilos entrados y salidos
End Sub

' *******  GESTION DE VENTAS  *********
' ******* INFORMES INCIDENCIAS ********

Public Sub SubmnC_GV_Incidencias_Click(Index As Integer)
    Select Case Index
        Case 1
            frmVtasListIncid.OpcionListado = 0  ' Informe de incidencias
        Case 2
            frmVtasListIncid.OpcionListado = 1  ' Informe de categorias
    End Select
    frmVtasListIncid.Show vbModal
End Sub

' *******  GESTION DE VENTAS  *********
' ******* INFORMES DATOS DE TRAZA ******** (SOLO NATURAL)

Public Sub SubmnC_GV_Traza_Click(Index As Integer)
    Select Case Index
        Case 1
            frmVtasListTraza.OpcionListado = 0  ' Informe de trazabilidad
            frmVtasListTraza.Show vbModal
    End Select
End Sub

' *******  GESTION DE VENTAS  *********
' *******    INFORMES         ********

Public Sub SubmnC_GV_InformesOficiales_Click()
    frmListOficiales.Show vbModal ' Informes Oficiales
End Sub

' *******  GESTION DE VENTAS  *********
' *******  FACTURAS VENTAS    ********

Public Sub SubmnC_GV_Facturas_Click(Index As Integer)
    Select Case Index
        Case 1: frmVtasFacturas.Show vbModal ' Facturas de venta
        Case 2: frmVtasReimpFact.Show vbModal  ' Reimpresión de facturas
        Case 3: AbrirListadoOfer 315 ' envio de facturas por email
        
        Case 4: AbrirListadoOfer 316 ' Facturacion Web/Electronica
        
        Case 5: frmVtasDiarioFact.Show vbModal ' Diario de Facturación
        Case 6: frmVtasAlbFact.Show vbModal ' Listado de Albaranes / Facturas
        
        Case 8: frmListCyC.Show vbModal ' Construc ("Facturas pendientes") 'Facturas pendientes (listado de cobros pendientes)
        Case 9: frmListCyCRiesgo.Show vbModal 'Construc ("Listado de Riesgo") 'Listado de Riesgo
        
        Case 11: Screen.MousePointer = vbHourglass 'frmVtasIntConta.Show vbModal ' Integracion contable
'                frmListado2.OptClientes = True
                frmListado2.OpcionListado = 223
                frmListado2.Show vbModal
                Screen.MousePointer = vbDefault
        Case 12: frmImpAridoc.Tipo = 1 ' Integracion de aridoc: Facturas de venta
                frmImpAridoc.Caption = "Importar Facturas a Aridoc"
                frmImpAridoc.Label4(16).Caption = "Fecha Factura"
                frmImpAridoc.Show vbModal
        Case 13: frmIntEdicom.Show vbModal  'Integracion edicom
        
        ' facturas a cuenta
        Case 15: frmVtasFacturasCta.Show vbModal ' Facturas a Cuenta
        
        Case 17: frmVtasAlbSocios.Show vbModal ' albaranes de venta a socios
        Case 18: frmVtasFactSocios.Show vbModal ' facturas de venta a socios
        Case 19: frmListado2.OpcionListado = 223
                 frmListado2.CadTag = "A"
                 frmListado2.Show vbModal
    End Select
End Sub

' *******  GESTION DE VENTAS  *********
' *******  FACTURAS TRANSPORTE ********

Public Sub SubmnC_GV_Portes_Click(Index As Integer)
    Select Case Index
        Case 1: frmVtasRecFactTrans.Show vbModal 'Recepcion de facturas de portes
        Case 2: frmVtasHcoFactTra.Show vbModal  'hco de facturas
        Case 4: DeTransporte = True
                AbrirListado2 (224) 'Informe de media gasto de transporte
        Case 6: frmVtasAlbTraPdtes.Show vbModal
        Case 7: frmVtasMediaTra.Show vbModal
    End Select
End Sub


' *******  GESTION DE VENTAS  *********
' *******  CONTROL DE COSTES  ********

Public Sub SubmnC_GV_CCostes_Click(Index As Integer)
    Select Case Index
        Case 1:  ' Areas
                AbrirFormularioAreasCC
        
        Case 2:  ' conceptos de costes
                frmCCManConcep.Show vbModal
            
        Case 3: 'frmCCManLineasConf.Show vbModal ' Lineas de confeccion
                AbrirFormularioLineasCC
                
        Case 4: ' rangos horarios
                frmCCHorario.Show vbModal
                
        Case 5: ' gastos/ingresos mensuales
                frmCCCostesMes.Show vbModal
                
        Case 7: 'procesar fichajes
                frmCCListados.OpcionListado = 2
                frmCCListados.Show vbModal
        
        Case 8: 'mantenimiento de fichajes
                frmCCFichajesTrab.Show vbModal
                
        Case 9: 'carga automatica de lineas de confeccion
                frmCCListados.OpcionListado = 1
                frmCCListados.Show vbModal
                
        Case 10: frmCCOrdenConfeccion.Show vbModal ' Ordenes de confeccion
        
        Case 11: frmCCCostesDiarios.Show vbModal ' Costes diarios
        
        Case 12: frmCCListados.OpcionListado = 0
                 frmCCListados.Show vbModal ' Informe de Costes diarios

        Case 13: frmCCListados.OpcionListado = 6
                 frmCCListados.Show vbModal ' busqueda de cadena en ficheros

    End Select
End Sub


' *******  INTEGRACION ANECOOP *********

Public Sub SubmnC_GV_Anecoop_Click(Index As Integer)

    Select Case Index
        Case 1: frmANECOOPTras.Show vbModal ' traspaso desde anecoop
        Case 2: frmANECOOPExped.Show vbModal ' expedientes de anecoop
        Case 3: frmANECOOPTrasFras.Show vbModal ' traspaso de facturas de anecoop
        Case 4: frmANECOOPTrasPagos.Show vbModal ' traspaso de pagos
    End Select

End Sub


'[Monica]18/02/2010: Ahora en recoleccion
'' *******  GESTION DE PRENOMINA *******
'Public Sub SubmnP_PreNominas_click(Index As Integer)
'    Select Case Index
'        Case 1: frmManSalarios.Show vbModal 'Mantenimiento de salarios
'        Case 2: frmManTraba.Show vbModal  'Mantenimiento de trabajadores
'        Case 4: frmManHoras.Show vbModal  'Entrada de Horas de trabajadores
'        Case 5: frmImpRecibos.Show vbModal 'Impresión de Recibos
'        Case 7: frmPagoRecibos.Show vbModal 'Pago de Recibos
'        Case 8: frmImpAridoc.tipo = 2 ' Integracion de aridoc: Recibos de Nóminas
'                frmImpAridoc.Caption = "Importar Recibos a Aridoc"
'                frmImpAridoc.Label4(16).Caption = "Fecha"
'                frmImpAridoc.Show vbModal 'vbModalConstruc("Integracion aridoc")
'        Case 10: frmInfHorasMes.Show vbModal ' Informe de Horas Mensual
'    End Select
'End Sub


' *******  UTILIDADES *********

Public Sub SubmnE_Util_Click(Index As Integer)
    Select Case Index
        Case 1: frmCaracteresMB.Show vbModal
        Case 3: frmCampPredet.Show vbModal
        Case 4: frmUtilidades.Show vbModal
        Case 6: frmLog.Show vbModal ' ver acciones
    End Select
End Sub

Public Sub AbrirListado2(numero As Integer)
    Screen.MousePointer = vbHourglass
    
    frmListado2.DeServicios = DeServicios
    frmListado2.OpcionListado = numero
    frmListado2.OptProve = (Not DeTransporte)
    DeTransporte = False
    
    frmListado2.Show vbModal
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub BloqueoMenusSegunContabilidad()
Dim b As Boolean

    b = (vParamAplic.NumeroConta <> 0)
    
'   la contabilizacion de facturas solo si hay contabilidad
    MDIppal.mnComGM_AlbCom(8).visible = MDIppal.mnComGM_AlbCom(8).visible And b
    MDIppal.mnComGM_AlbCom(8).Enabled = MDIppal.mnComGM_AlbCom(8).visible And b
    
'   los tipos de iva unicamente los tenemos en la bd agro si no hay contabilidad
    MDIppal.mnComG_Admon(10).visible = MDIppal.mnComG_Admon(10).visible And Not b
    MDIppal.mnComG_Admon(10).Enabled = MDIppal.mnComG_Admon(10).visible And Not b
    
End Sub

Public Sub BloqueoMenusSegunNivelUsuario()

    MDIppal.mnComGM_Gral(3).visible = MDIppal.mnComGM_Gral(3).visible And (vUsu.Nivel <= 1)
    MDIppal.mnComGM_Gral(3).Enabled = MDIppal.mnComGM_Gral(3).visible And (vUsu.Nivel <= 1)
    
    
    MDIppal.mnP_Generales(9).visible = MDIppal.mnP_Generales(9).visible And (vUsu.Nivel = 0)
    MDIppal.mnP_Generales(9).Enabled = MDIppal.mnP_Generales(9).visible And (vUsu.Nivel = 0)
    
    MDIppal.mnE_Util(4).visible = (vUsu.Nivel = 0)
    MDIppal.mnE_Util(4).Enabled = (vUsu.Nivel = 0)
    
    MDIppal.mnComGV_Costes(6).visible = (vUsu.Nivel <= 1)
    MDIppal.mnComGV_Costes(6).Enabled = (vUsu.Nivel <= 1)

    MDIppal.mnComGV_Costes(13).visible = (vUsu.Nivel = 0)
    MDIppal.mnComGV_Costes(13).Enabled = (vUsu.Nivel = 0)

    MDIppal.mnE_Util(6).visible = (vUsu.Login = "root")
    MDIppal.mnE_Util(6).Enabled = (vUsu.Login = "root")

End Sub

Public Sub BloqueoMenusSegunCooperativa()

    MDIppal.mnComerGV_Traza(1).visible = (vParamAplic.Cooperativa = 9)
    MDIppal.mnComerGV_Traza(1).Enabled = (vParamAplic.Cooperativa = 9)

End Sub



Private Sub AbrirFormularioAreasCC()
    
    Set frmBas = New frmBasico
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|800|;S|txtAux(1)|T|Descripción|3930|;"
    frmBas.CadenaConsulta = "SELECT ccareas.codarea, ccareas.nomarea "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM ccareas "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "Código|N|N|0|9999|ccareas|codarea|0000|S|"
    frmBas.Tag2 = "Descripción|T|N|||ccareas|nomarea|||"
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 50
    frmBas.Tabla = "ccareas"
    frmBas.CampoCP = "codarea"
    frmBas.Report = "rManCCAreas.rpt"
    frmBas.Caption = "Áreas"
    frmBas.Show vbModal
    
    Set frmBas = Nothing

End Sub


Private Sub AbrirFormularioZonasCC()
    
    Set frmBas = New frmBasico
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|800|;S|txtAux(1)|T|Descripción|3930|;"
    frmBas.CadenaConsulta = "SELECT cczonas.codzona, cczonas.nomzona "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM cczonas "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "Código|N|N|0|9999|cczonas|codzona|0000|S|"
    frmBas.Tag2 = "Descripción|T|N|||cczonas|nomzona|||"
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 30
    frmBas.Tabla = "cczonas"
    frmBas.CampoCP = "codzona"
    frmBas.Report = "rManCCZonas.rpt"
    frmBas.Caption = "Zonas"
    frmBas.Show vbModal
    
    Set frmBas = Nothing

End Sub

Private Sub AbrirFormularioLineasCC()
    
    Set frmBas = New frmBasico
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|800|;S|txtAux(1)|T|Descripción|3930|;"
    frmBas.CadenaConsulta = "SELECT cclinconf.codlinconf, cclinconf.nomlinconf "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM cclinconf "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "Código|N|N|0|9999|cclinconf|codlinconf|00|S|"
    frmBas.Tag2 = "Descripción|T|N|||cclinconf|nomlinconf|||"
    frmBas.Maxlen1 = 2
    frmBas.Maxlen2 = 40
    frmBas.Tabla = "cclinconf"
    frmBas.CampoCP = "codlinconf"
    frmBas.Report = "rManCCLineasConf.rpt"
    frmBas.Caption = "Lineas de Confección"
    frmBas.Show vbModal
    
    Set frmBas = Nothing

End Sub


