VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIppal 
   BackColor       =   &H8000000C&
   Caption         =   "Ariagro - Gestión Comercial"
   ClientHeight    =   7860
   ClientLeft      =   225
   ClientTop       =   1170
   ClientWidth     =   11160
   Icon            =   "MDIppal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   23
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clientes"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Variedades"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Forfaits"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Proveedores"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Artículos"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   4
            Object.Width           =   200
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedidos Compra"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Albaranes Compra"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Compras"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Recepción Facturas"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   4
            Object.Width           =   200
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Palets"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedidos Venta"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Albaranes Venta"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Albarán Envases "
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   4
            Object.Width           =   200
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Introducción Factura"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Transporte"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Informes"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambio de Campaña"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   1
      Top             =   7275
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3176
            MinWidth        =   3176
            Picture         =   "MDIppal.frx":6852
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2471
            MinWidth        =   2471
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3096
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5997
            MinWidth        =   5997
            Picture         =   "MDIppal.frx":7132
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "16:22"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnParametros 
      Caption         =   "&Datos Básicos"
      Index           =   1
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Datos de Empresa"
         Index           =   1
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Parámetros"
         Index           =   2
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Tipos de Movimiento"
         Index           =   3
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Tipos de Documentos"
         Index           =   4
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Usuarios"
         Index           =   6
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Cambio de Campaña"
         Index           =   7
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "Creación Nueva Campaña"
         Index           =   8
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "Traspaso C&hivato"
         Index           =   9
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Salir"
         Index           =   11
      End
   End
   Begin VB.Menu mnComerGen 
      Caption         =   "Datos &Generales"
      Begin VB.Menu mnComerAdm 
         Caption         =   "&Administracion"
         Begin VB.Menu mnComG_Admon 
            Caption         =   "&Países"
            Index           =   1
         End
         Begin VB.Menu mnComG_Admon 
            Caption         =   "&Tipos de Mercado"
            Index           =   2
         End
         Begin VB.Menu mnComG_Admon 
            Caption         =   "&Cadenas"
            Index           =   3
         End
         Begin VB.Menu mnComG_Admon 
            Caption         =   "&Agencias de Transporte"
            Index           =   4
         End
         Begin VB.Menu mnComG_Admon 
            Caption         =   "&Formas de Pago"
            Index           =   5
         End
         Begin VB.Menu mnComG_Admon 
            Caption         =   "&Clientes"
            Index           =   6
         End
         Begin VB.Menu mnComG_Admon 
            Caption         =   "&Bancos"
            Index           =   7
         End
         Begin VB.Menu mnComG_Admon 
            Caption         =   "&Tipos de Cartas"
            Index           =   8
         End
         Begin VB.Menu mnComG_Admon 
            Caption         =   "&Incidencias"
            Index           =   9
         End
         Begin VB.Menu mnComG_Admon 
            Caption         =   "Tipos de &Iva"
            Index           =   10
         End
      End
      Begin VB.Menu mnComerPro 
         Caption         =   "&Productos"
         Begin VB.Menu mnComG_Produc 
            Caption         =   "&Grupos"
            Index           =   1
         End
         Begin VB.Menu mnComG_Produc 
            Caption         =   "&Productos"
            Index           =   2
         End
         Begin VB.Menu mnComG_Produc 
            Caption         =   "C&lases"
            Index           =   3
         End
         Begin VB.Menu mnComG_Produc 
            Caption         =   "&Variedades"
            Index           =   4
         End
         Begin VB.Menu mnComG_Produc 
            Caption         =   "&Calibres"
            Index           =   5
         End
         Begin VB.Menu mnComG_Produc 
            Caption         =   "&Marcas"
            Index           =   6
         End
         Begin VB.Menu mnComG_Produc 
            Caption         =   "Códigos &Ean"
            Index           =   7
         End
      End
      Begin VB.Menu mnComerConfec 
         Caption         =   "&Confeccion"
         Begin VB.Menu mnComG_Confec 
            Caption         =   "&Tipos de Envases"
            Index           =   1
         End
         Begin VB.Menu mnComG_Confec 
            Caption         =   "&Capacidad Envases"
            Index           =   2
         End
         Begin VB.Menu mnComG_Confec 
            Caption         =   "&Medidas Envases"
            Index           =   3
         End
         Begin VB.Menu mnComG_Confec 
            Caption         =   "&Forma de Confección"
            Index           =   4
         End
         Begin VB.Menu mnComG_Confec 
            Caption         =   "&Presentación"
            Index           =   5
         End
         Begin VB.Menu mnComG_Confec 
            Caption         =   "Paletización"
            Index           =   6
         End
         Begin VB.Menu mnComG_Confec 
            Caption         =   "&Descripciones Costes Confección"
            Index           =   7
         End
         Begin VB.Menu mnComG_Confec 
            Caption         =   "&Forfaits"
            Index           =   8
         End
      End
   End
   Begin VB.Menu mnGesMater 
      Caption         =   "Gestión &Materiales"
      Begin VB.Menu mnComerGM_Gral 
         Caption         =   "&Datos Generales"
         Begin VB.Menu mnComGM_Gral 
            Caption         =   "Almacenes Propios"
            Index           =   1
         End
         Begin VB.Menu mnComGM_Gral 
            Caption         =   "&Tipos de Unidad"
            Index           =   2
         End
         Begin VB.Menu mnComGM_Gral 
            Caption         =   "Tipos de Artículos"
            Index           =   3
         End
         Begin VB.Menu mnComGM_Gral 
            Caption         =   "&Familias"
            Index           =   4
         End
         Begin VB.Menu mnComGM_Gral 
            Caption         =   "&Articulos"
            Index           =   5
         End
         Begin VB.Menu mnComGM_Gral 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnComGM_Gral 
            Caption         =   "&Proveedores"
            Index           =   7
         End
         Begin VB.Menu mnComGM_Gral 
            Caption         =   "Proveedores &Varios"
            Index           =   8
         End
         Begin VB.Menu mnComGM_Gral 
            Caption         =   "&Direcciones"
            Index           =   9
         End
         Begin VB.Menu mnComGM_Gral 
            Caption         =   "-"
            Index           =   10
         End
         Begin VB.Menu mnComGM_Gral 
            Caption         =   "Precios Proveedor"
            Index           =   11
         End
      End
      Begin VB.Menu mnComerGM_InfVa 
         Caption         =   "&Informes Varios"
         Begin VB.Menu mnComGM_InfVa 
            Caption         =   "&Proveedores"
            Index           =   1
         End
         Begin VB.Menu mnComGM_InfVa 
            Caption         =   "&Etiquetas de Proveedores"
            Index           =   2
         End
         Begin VB.Menu mnComGM_InfVa 
            Caption         =   "&Cartas a Proveedores"
            Index           =   3
         End
      End
      Begin VB.Menu mnComerGM 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnComerGM_PedCom 
         Caption         =   "&Pedidos Compras"
         Begin VB.Menu mnComGM_PedCom 
            Caption         =   "&Mant.Pedidos Proveedor"
            Index           =   1
         End
         Begin VB.Menu mnComGM_PedCom 
            Caption         =   "&Histórico Pedidos Anulados"
            Index           =   2
         End
         Begin VB.Menu mnComGM_PedCom 
            Caption         =   "&List. Material pendiente de recibir"
            Index           =   3
         End
      End
      Begin VB.Menu mnComerGM_AlbCom 
         Caption         =   "&Albaranes Compras"
         Begin VB.Menu mnComGM_AlbCom 
            Caption         =   "&Mant. Albaranes Proveedor"
            Index           =   1
         End
         Begin VB.Menu mnComGM_AlbCom 
            Caption         =   "&Histórico Albaranes Anulados"
            Index           =   2
         End
         Begin VB.Menu mnComGM_AlbCom 
            Caption         =   "&List. Pendiente de facturar"
            Index           =   3
         End
         Begin VB.Menu mnComGM_AlbCom 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnComGM_AlbCom 
            Caption         =   "&Recepción Facturas"
            Index           =   5
         End
         Begin VB.Menu mnComGM_AlbCom 
            Caption         =   "&Histórico Albarán/Factura"
            Index           =   6
         End
         Begin VB.Menu mnComGM_AlbCom 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnComGM_AlbCom 
            Caption         =   "&Contabilizar Facturas"
            Index           =   8
         End
      End
      Begin VB.Menu mnComerGM_EstCom 
         Caption         =   "&Estadísticas Compras"
         Begin VB.Menu mnComGM_EstCom 
            Caption         =   "Compras por &Proveedor"
            Index           =   1
         End
         Begin VB.Menu mnComGM_EstCom 
            Caption         =   "Compras por &Familia/Artículo"
            Index           =   2
         End
      End
      Begin VB.Menu mnComerGM1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnComerGM_MovAlm 
         Caption         =   "&Movimientos de Almacen"
         Begin VB.Menu mnComGM_MovAlm 
            Caption         =   "&Traspaso de Almacenes"
            Index           =   1
         End
         Begin VB.Menu mnComGM_MovAlm 
            Caption         =   "&Histórico Traspaso Almacenes"
            Index           =   2
         End
         Begin VB.Menu mnComGM_MovAlm 
            Caption         =   "&Movimientos Almacen"
            Index           =   3
         End
         Begin VB.Menu mnComGM_MovAlm 
            Caption         =   "Histórico Movimientos &Almacén"
            Index           =   4
         End
         Begin VB.Menu mnComGM_MovAlm 
            Caption         =   "Movimientos &Servicios Varios"
            Index           =   5
         End
         Begin VB.Menu mnComGM_MovAlm 
            Caption         =   "Histórico Servicios &Varios"
            Index           =   6
         End
         Begin VB.Menu mnComGM_MovAlm 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnComGM_MovAlm 
            Caption         =   "&Envases Retornables"
            Index           =   8
         End
         Begin VB.Menu mnComGM_MovAlm 
            Caption         =   "&Grabación Fichero Cheps"
            Index           =   9
         End
      End
      Begin VB.Menu mnComerGM_Consul 
         Caption         =   "&Consultas"
         Begin VB.Menu mnComGM_Consultas 
            Caption         =   "&Movimientos de Artículos"
            Index           =   1
         End
         Begin VB.Menu mnComGM_Consultas 
            Caption         =   "&Listado de Movimientos"
            Index           =   2
         End
         Begin VB.Menu mnComGM_Consultas 
            Caption         =   "&Listado Valoración Stock"
            Index           =   3
         End
         Begin VB.Menu mnComGM_Consultas 
            Caption         =   "&Informe Stocks Máximos_Mínimos"
            Index           =   4
         End
         Begin VB.Menu mnComGM_Consultas 
            Caption         =   "Informe &Stocks a una Fecha"
            Index           =   5
         End
         Begin VB.Menu mnComGM_Consultas 
            Caption         =   "Movimientos &Articulos por Familia"
            Index           =   6
         End
         Begin VB.Menu mnComGM_Consultas 
            Caption         =   "&Movimientos Artículos por Socio/Cliente"
            Index           =   7
         End
      End
      Begin VB.Menu mnComerGM_Invent 
         Caption         =   "&Inventario"
         Begin VB.Menu mnComGM_Invent 
            Caption         =   "&Toma de Inventario"
            Index           =   1
         End
         Begin VB.Menu mnComGM_Invent 
            Caption         =   "&Entrada Existencia Real"
            Index           =   2
         End
         Begin VB.Menu mnComGM_Invent 
            Caption         =   "&Listado de Diferencias"
            Index           =   3
         End
         Begin VB.Menu mnComGM_Invent 
            Caption         =   "&Actualizar Diferencias"
            Index           =   4
         End
         Begin VB.Menu mnComGM_Invent 
            Caption         =   "&Valoración Stocks inventariados"
            Index           =   5
         End
         Begin VB.Menu mnComGM_Invent 
            Caption         =   "&Histórico inventario"
            Index           =   6
         End
         Begin VB.Menu mnComGM_Invent 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnComGM_Invent 
            Caption         =   "Recálculo Precio Medio Ponderado"
            Index           =   8
         End
      End
   End
   Begin VB.Menu mnGesVtas 
      Caption         =   "&Palets-Pedidos-Albaranes"
      Begin VB.Menu mnComGV_Palets 
         Caption         =   "&Carga Automática Palets"
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnComGV_PalPed 
         Caption         =   "&Palets"
         Index           =   1
      End
      Begin VB.Menu mnComGV_PalPed 
         Caption         =   "&Informe Palets Confeccionados"
         Index           =   2
      End
      Begin VB.Menu mnComGV_PalPed 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnComGV_PalPed 
         Caption         =   "Pe&didos"
         Index           =   4
      End
      Begin VB.Menu mnComGV_PalPed 
         Caption         =   "&Planning Diario"
         Index           =   5
      End
      Begin VB.Menu mnComGV_PalPed 
         Caption         =   "&Orden de Confección"
         Index           =   6
      End
      Begin VB.Menu mnComGV_PalPed 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnComGV_PalPed 
         Caption         =   "&Albaranes"
         Index           =   8
      End
      Begin VB.Menu mnComGV_PalPed 
         Caption         =   "&Revalorar costes Albaranes"
         Index           =   9
      End
      Begin VB.Menu mnComGV_PalPed 
         Caption         =   "Recalcular &Costes Reales"
         Index           =   10
      End
      Begin VB.Menu mnComGV_PalPed 
         Caption         =   "Albaranes &pdtes facturar"
         Index           =   11
      End
      Begin VB.Menu mnComGV_PalPed 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnComGV_PalPed 
         Caption         =   "Integración Aridoc"
         Index           =   13
      End
      Begin VB.Menu mnComGV_PalPed 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnComGV_PalPed 
         Caption         =   "Traspaso de Trazabilidad"
         Index           =   15
      End
      Begin VB.Menu mnComerGV1 
         Caption         =   "-"
      End
      Begin VB.Menu mnComerGV_AlbEnv 
         Caption         =   "&Albaranes Envases"
      End
      Begin VB.Menu mnComerGV_Factu 
         Caption         =   "&Facturación"
      End
      Begin VB.Menu mnComerGV3 
         Caption         =   "-"
      End
      Begin VB.Menu mnComerGV_Inf 
         Caption         =   "&Informes"
         Begin VB.Menu mnComerGV_Informes 
            Caption         =   "&Informes Ventas"
         End
         Begin VB.Menu mnComerGV_Rdto 
            Caption         =   "&Rendimiento por Variedad"
         End
         Begin VB.Menu mnComerGV_InfGra 
            Caption         =   "Informes &Gráficos"
         End
         Begin VB.Menu mnComerGV_Intrastat 
            Caption         =   "Relación Intrastat"
         End
         Begin VB.Menu mnComerGV_Diferencias 
            Caption         =   "Diferencias Kilos E/S"
         End
         Begin VB.Menu mnComerGV_Incidencias 
            Caption         =   "&Informe de Incidencias"
            Index           =   1
         End
         Begin VB.Menu mnComerGV_Incidencias 
            Caption         =   "&Informe de Categorias"
            Index           =   2
         End
         Begin VB.Menu mnComerGV_Traza 
            Caption         =   "&Informe de Trazabilidad"
            Index           =   1
         End
         Begin VB.Menu mnComerGV4 
            Caption         =   "-"
         End
         Begin VB.Menu mnComerGV_InfOficial 
            Caption         =   "Informes &Oficiales"
         End
      End
   End
   Begin VB.Menu mnComerGV_Facturas 
      Caption         =   "&Facturas"
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "&Introducción Factura"
         Index           =   1
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "&Reimpresión Facturas"
         Index           =   2
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "Enviar Facturas por email"
         Index           =   3
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "Facturación &Web/Electrónica"
         Index           =   4
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "Diario de Facturación"
         Index           =   5
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "List.Albaranes/Facturas"
         Index           =   6
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "&Facturas Pendientes"
         Index           =   8
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "&Listado de Riesgo"
         Index           =   9
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "Integración &Contable"
         Index           =   11
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "Integración &Aridoc"
         Index           =   12
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "Integración &Edicom"
         Index           =   13
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "Facturas a Cuen&ta"
         Index           =   15
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "Albaranes Venta Socio"
         Index           =   17
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "Facturas Ve&nta Socio"
         Index           =   18
      End
      Begin VB.Menu mnComGV_Facturas 
         Caption         =   "Inte&gración Contable"
         Index           =   19
      End
   End
   Begin VB.Menu mnComerGV_Portes 
      Caption         =   "Facturas &Transporte/Comisión"
      Begin VB.Menu mnComGV_Portes 
         Caption         =   "&Recepción Facturas"
         Index           =   1
      End
      Begin VB.Menu mnComGV_Portes 
         Caption         =   "&Histórico Albarán/Factura"
         Index           =   2
      End
      Begin VB.Menu mnComGV_Portes 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnComGV_Portes 
         Caption         =   "&Integración Contable"
         Index           =   4
      End
      Begin VB.Menu mnComGV_Portes 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnComGV_Portes 
         Caption         =   "&Albaranes Pdtes.Facturar"
         Index           =   6
      End
      Begin VB.Menu mnComGV_Portes 
         Caption         =   "&Informe Coste de Transporte"
         Index           =   7
      End
   End
   Begin VB.Menu mnComerGV_Costes 
      Caption         =   "Control Costes"
      Begin VB.Menu mnComGV_Costes 
         Caption         =   "&Áreas Generales"
         Index           =   1
      End
      Begin VB.Menu mnComGV_Costes 
         Caption         =   "&Conceptos de Costes"
         Index           =   2
      End
      Begin VB.Menu mnComGV_Costes 
         Caption         =   "&Líneas de Confección"
         Index           =   3
      End
      Begin VB.Menu mnComGV_Costes 
         Caption         =   "&Rangos Horarios"
         Index           =   4
      End
      Begin VB.Menu mnComGV_Costes 
         Caption         =   "&Gastos/Ingresos Mensuales"
         Index           =   5
      End
      Begin VB.Menu mnComGV_Costes 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnComGV_Costes 
         Caption         =   "&Procesar Fichajes"
         Index           =   7
      End
      Begin VB.Menu mnComGV_Costes 
         Caption         =   "&Mantenimiento Fichajes"
         Index           =   8
      End
      Begin VB.Menu mnComGV_Costes 
         Caption         =   "Carga Ordenes Confección"
         Index           =   9
      End
      Begin VB.Menu mnComGV_Costes 
         Caption         =   "&Orden de Confección"
         Enabled         =   0   'False
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnComGV_Costes 
         Caption         =   "Costes &Diarios"
         Enabled         =   0   'False
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnComGV_Costes 
         Caption         =   "&Informe de Costes"
         Index           =   12
      End
      Begin VB.Menu mnComGV_Costes 
         Caption         =   "Búsqueda en &Ficheros"
         Index           =   13
      End
   End
   Begin VB.Menu mnIntAnecoop 
      Caption         =   "&Integración Anecoop"
      Begin VB.Menu mnAnecoop 
         Caption         =   "&Traspaso Anecoop"
         Index           =   1
      End
      Begin VB.Menu mnAnecoop 
         Caption         =   "&Expedientes"
         Index           =   2
      End
      Begin VB.Menu mnAnecoop 
         Caption         =   "Traspaso de &Facturas"
         Index           =   3
      End
      Begin VB.Menu mnAnecoop 
         Caption         =   "Traspaso de &Pagos"
         Index           =   4
      End
   End
   Begin VB.Menu mnUtil 
      Caption         =   "&Utilidades"
      WindowList      =   -1  'True
      Begin VB.Menu mnE_Util 
         Caption         =   "Revisión de caracteres en Multibase"
         Index           =   1
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "&Cambio Campaña Actual"
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "&Insercion facturas_calib"
         Index           =   4
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Acciones Realizadas"
         Index           =   6
      End
   End
   Begin VB.Menu mnSoporte 
      Caption         =   "&Soporte"
      Begin VB.Menu mnE_Soporte1 
         Caption         =   "&Web Soporte"
      End
      Begin VB.Menu mnp_Barra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnE_Soporte2 
         Caption         =   "&Acerca de"
      End
   End
End
Attribute VB_Name = "MDIppal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private PrimeraVez As Boolean
Dim TieneEditorDeMenus As Boolean

Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal op As Integer, ByVal tam As Integer)
    Dim i As Integer
    Dim tRes As ResType, iCount As Integer
        
    opcio = op
    tamany = tam
    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)

    If ghmodule = 0 Then
        MsgBox "Invalid library file.", vbCritical
        Exit Sub
    End If
        
    For tRes = RT_FIRST To RT_LAST
        DoEvents
        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
    Next
    FreeLibrary ghmodule
             
End Sub




Private Sub MDIForm_Activate()
'Dim cad As String

    If PrimeraVez Then
        PrimeraVez = False
'        frmMensaje.pTitulo = "Últimas modificaciones.         14/03/06"
'
''        cad = cad & "-----------------------------------------------------------------------------------------------" & vbCrLf
''        cad = cad & "Para actualizar el estado de un presupuesto desde la pantalla de ventas sin "
''        cad = cad & "entrar en la pantalla de modificación de presupuesto, seleccionar la línea del "
''        cad = cad & "presupuesto a modificar, pulsar botón izquierdo del ratón, se despliega un menu "
''        cad = cad & "con los posibles estados y se selecciona el nuevo estado." & vbCrLf & vbCrLf
'
''        cad = cad & "- Imprimir informes de subcontratación." & vbCrLf
''        cad = cad & "- Ventas pendientes." & vbCrLf
''        cad = cad & "-------------------------------------------------------------------------" & vbCrLf & vbCrLf
'
'        cad = cad & "- Mantenimiento de No Conformidades y lineas de acciones y reclamaciones." & vbCrLf & vbCrLf
'        cad = cad & "- Informes:" & vbCrLf
'        cad = cad & "     Comunicación con cliente." & vbCrLf
'        cad = cad & "     Confirmación de servicios." & vbCrLf
'        cad = cad & "     No conformidad." & vbCrLf
'        cad = cad & "     Reclamación." & vbCrLf & vbCrLf
'
'
'        frmMensaje.pValor = cad
'        frmMensaje.Show vbModal
    End If
End Sub

Private Sub MDIForm_Load()
Dim cad As String

    PrimeraVez = True
    CargarImagen
    PonerDatosFormulario

    
    If vParam Is Nothing Then
        Caption = "AriAgro - Gestión Comercial" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & " FALTA CONFIGURAR"
    Else
        Caption = "AriAgro - Gestión Comercial" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  Empresa: " & vParam.NombreEmpresa & cad & _
                  " - Campaña: " & vParam.FecIniCam & " - " & vParam.FecFinCam & "   -  Usuario: " & vUsu.Nombre
    End If

    ' *** per als iconos XP ***
'--monica: quitamos los iconos de 32
'    GetIconsFromLibrary App.Path & "\iconos.dll", 1, 32
'
    GetIconsFromLibrary App.path & "\iconos.dll", 1, 24
    GetIconsFromLibrary App.path & "\iconos_BN.dll", 2, 24
    GetIconsFromLibrary App.path & "\iconos_OM.dll", 3, 24
    
    GetIconsFromLibrary App.path & "\iconosAriagro.dll", 4, 24
    
 
  
    'CARGAR LA TOOLBAR DEL FORM PRINCIPAL
    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListPpal

        .Buttons(1).Image = 2   'Clientes
        .Buttons(2).Image = 12   'Variedades
        .Buttons(3).Image = 13   'Forfaits
        .Buttons(4).Image = 18   'Proveedores
        .Buttons(5).Image = 19   'Artículos
        'el 6 separador
        .Buttons(7).Image = 15   'Pedidos
        .Buttons(8).Image = 10  'Albaranes
        .Buttons(9).Image = 22  'Facturas compra
        .Buttons(10).Image = 24  'Recepcion Facturas compra
        
        'el 11 separador
        .Buttons(12).Image = 14   'Palets
        .Buttons(13).Image = 4   'Pedidos
        .Buttons(14).Image = 25  'Albaranes de Venta
        .Buttons(15).Image = 27  ' albaranes de envases
        'el 16 separador
        .Buttons(17).Image = 23   ' Introduccion de Factura
        .Buttons(18).Image = 7   'Facturas Transporte
        .Buttons(19).Image = 21   'Informes
        'el 20 separador
        .Buttons(21).Image = 26   'Cambio de campaña
        'el 22 separador
        .Buttons(23).Image = 1   'Salir
    End With
    
    
    
    GetIconsFromLibrary App.path & "\iconos.dll", 1, 16
    GetIconsFromLibrary App.path & "\iconos_BN.dll", 2, 16
    GetIconsFromLibrary App.path & "\iconos_OM.dll", 3, 16

    LeerEditorMenus

    PonerDatosFormulario
    
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    AccionesCerrar
    End
End Sub

Private Sub mnAnecoop_Click(Index As Integer)
    SubmnC_GV_Anecoop_Click (Index)
End Sub

Private Sub mnComerGV_AlbEnv_Click()
    SubmnC_GV_AlbEnv_Click
End Sub

Private Sub mnComerGV_Diferencias_Click()
    SubmnC_GV_Diferencias_Click
End Sub

Private Sub mnComerGV_Factu_Click()
    SubmnC_GV_FactAlbEnv_Click
End Sub

Private Sub mnComerGV_Incidencias_Click(Index As Integer)
    SubmnC_GV_Incidencias_Click (Index)
End Sub

Private Sub mnComerGV_InfGra_Click()
    SubmnC_GV_InfGra_Click
End Sub

Private Sub mnComerGV_InfOficial_Click()
    SubmnC_GV_InformesOficiales_Click
End Sub

Private Sub mnComerGV_Informes_Click()
    SubmnC_GV_Informes_Click
End Sub

Private Sub mnComerGV_Intrastat_Click()
    SubmnC_GV_Intrastat_Click
End Sub

Private Sub mnComerGV_Rdto_Click()
    SubmnC_GV_Rdto_Click
End Sub

Private Sub mnComerGV_Traza_Click(Index As Integer)
    SubmnC_GV_Traza_Click (Index)
End Sub

Private Sub mnComG_Admon_Click(Index As Integer)
    SubmnC_ComercialG_Admon_Click (Index)
End Sub

Private Sub mnComG_Confec_Click(Index As Integer)
    SubmnC_ComercialG_Confec_Click (Index)
End Sub

Private Sub mnComG_Produc_Click(Index As Integer)
    SubmnC_ComercialG_Produc_Click (Index)
End Sub

Private Sub mnComGM_AlbCom_Click(Index As Integer)
    SubmnC_GM_AlbCom_Click (Index)
End Sub

Private Sub mnComGM_EstCom_Click(Index As Integer)
    SubmnC_GM_EstCom_Click (Index)
End Sub

Private Sub mnComGM_InfVa_Click(Index As Integer)
    SubmnC_GM_InfVa_Click (Index)
End Sub

Private Sub mnComGM_Invent_Click(Index As Integer)
    SubmnC_GM_Inven_Click (Index)
End Sub

Private Sub mnComGM_MovAlm_Click(Index As Integer)
    SubmnC_GM_MovAlm_Click (Index)
End Sub

Private Sub mnComGM_Consultas_Click(Index As Integer)
    SubmnC_GM_Consul_Click (Index)
End Sub

Private Sub mnComGM_PedCom_Click(Index As Integer)
    SubmnC_GM_PedCom_Click (Index)
End Sub

Private Sub mnComGV_Costes_Click(Index As Integer)
    SubmnC_GV_CCostes_Click (Index)
End Sub

Private Sub mnComGV_Facturas_Click(Index As Integer)
    SubmnC_GV_Facturas_Click (Index)
End Sub

Private Sub mnComGV_Palets_Click(Index As Integer)
    SubmnC_GV_Palets_Click (Index)
End Sub

Private Sub mnComGV_PalPed_Click(Index As Integer)
    SubmnC_GV_PalPed_Click (Index)
End Sub

Private Sub mnComGV_Portes_Click(Index As Integer)
    SubmnC_GV_Portes_Click (Index)
End Sub

Private Sub mnE_Soporte1_Click()
    Screen.MousePointer = vbHourglass
    LanzaHome "websoporte"
    espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnE_Util_Click(Index As Integer)
    SubmnE_Util_Click (Index)
End Sub


Private Sub mnE_Soporte2_Click()
    frmMensaje.OpcionMensaje = 6
    frmMensaje.Show vbModal
End Sub

Private Sub mnComGM_Gral_Click(Index As Integer)
    SubmnC_GM_Gral_Click (Index)
End Sub

'Private Sub mnGesVtas_Click(index As Integer)
'    SubmnC_GV_Gral_Click (index)
'End Sub

Private Sub mnP_Generales_Click(Index As Integer)
    If Index = 7 Then
        mnCambioEmpresa_Click
    Else
        SubmnP_Generales_Click (Index)
    End If
End Sub

Private Sub BotonSalir()
    Unload frmPpal
    Unload Me
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Clientes
            SubmnC_ComercialG_Admon_Click (6)
        Case 2 'Variedades
            SubmnC_ComercialG_Produc_Click (4)
        Case 3 'Forfaits
            SubmnC_ComercialG_Confec_Click (8)
        Case 4 'Proveedores
            SubmnC_GM_Gral_Click (7)
        Case 5 'Articulos
            SubmnC_GM_Gral_Click (5)
        Case 7 'Pedidos
            SubmnC_GM_PedCom_Click (1)
        Case 8 'Albaranes
            SubmnC_GM_AlbCom_Click (1)
        Case 9 'Facturas compras
            SubmnC_GM_AlbCom_Click (6)
        Case 10 'Recepcion Facturas compras
            SubmnC_GM_AlbCom_Click (5)
        Case 12 'Palets
            SubmnC_GV_PalPed_Click (1)
        Case 13 'Pedidos ventas
            SubmnC_GV_PalPed_Click (4)
        Case 14 'albaranes ventas
            SubmnC_GV_PalPed_Click (8)
        Case 15 ' albaran de envases
            SubmnC_GV_AlbEnv_Click
        Case 17 'Facturacion
            SubmnC_GV_Facturas_Click (1)
        Case 18 'Facturas transporte
            SubmnC_GV_Portes_Click (2)
        Case 19 'Informes
            SubmnC_GV_Informes_Click
        Case 21 'Cambio de Campaña
            mnCambioEmpresa_Click
        Case 23 ' Salir
            MDIForm_Unload 0
    End Select
End Sub

' ### [Monica] 05/09/2006
Private Sub PonerDatosFormulario()
Dim Config As Boolean

    Config = (vEmpresa Is Nothing) 'Or (vParamAplic Is Nothing)
    
    If Not Config Then HabilitarSoloPrametros_o_Empresas True

    'FijarConerrores
    CadenaDesdeOtroForm = ""

    'Poner datos visible del form
    PonerDatosVisiblesForm
    
    'Habilitar/Deshabilitar entradas del menu segun el nivel de usuario
    PonerMenusNivelUsuario

    'Habilitar
    If Config Then HabilitarSoloPrametros_o_Empresas False
    'Panel con el nombre de la empresa
'    If Not vEmpresa Is Nothing Then
'        Me.StatusBar1.Panels(2).Text = "Empresa:   " & vEmpresa.nomempre & "               Código: " & vEmpresa.codempre
'    Else
'        Me.StatusBar1.Panels(2).Text = "Falta configurar"
'    End If


    'Si tiene editor de menus
    If TieneEditorDeMenus Then
        PoneMenusDelEditor
    End If
    
    PoneBotonesdelToolbar
    
    BloqueoMenusSegunContabilidad
    
    BloqueoMenusSegunNivelUsuario
    
    BloqueoMenusSegunCooperativa
    
    '[Monica]02/11/2011: bloqueamos cuando no está activo el parametro
    Me.mnComerGV_Costes.Enabled = vParamAplic.HayCCostes
    Me.mnComerGV_Costes.visible = vParamAplic.HayCCostes
    
    
    'Bloqueo si hayaridoc
    Me.mnComGV_PalPed(13).Enabled = vParamAplic.HayAridoc
    Me.mnComGV_Facturas(12).Enabled = vParamAplic.HayAridoc
    
    
    
    '[Monica]31/03/2015: integracion anecoop
    Me.mnIntAnecoop.Enabled = vParamAplic.HayAnecoop
    Me.mnIntAnecoop.visible = vParamAplic.HayAnecoop
    
    '[Monica]27/12/2012: Bloqueo de la facturacion electronica si no estamos en ariagro
    '[Monica]07/03/2013: quito la condicion de que solo cuando es campaña actual, puede ser en cualquier campaña.
    Me.mnComGV_Facturas(4).Enabled = (vParamAplic.PathFacturaE <> "") ' And EsCampanyaActual(vEmpresa.BDAriagro)
    
End Sub

' ### [Monica] 05/09/2006
Private Sub HabilitarSoloPrametros_o_Empresas(Habilitar As Boolean)
Dim T As Control
Dim cad As String

    On Error Resume Next
    For Each T In Me
        cad = T.Name
        If Mid(T.Name, 1, 2) = "mn" Then
            'If LCase(Mid(T.Name, 1, 8)) <> "mn_b" Then
                T.Enabled = Habilitar
            'End If
        End If
    Next
    
    Me.Toolbar1.Enabled = Habilitar
    Me.Toolbar1.visible = Habilitar
    Me.mnParametros(1).Enabled = True
    Me.mnP_Generales(1).Enabled = True
    Me.mnP_Generales(2).Enabled = True
    Me.mnP_Generales(6).Enabled = True
    Me.mnP_Generales(17).Enabled = True
    
'    Me.mnCambioEmpresa.Enabled = True
End Sub


' ### [Monica] 07/11/2006
' añadida esta parte para la personalizacion de menus

Private Sub LeerEditorMenus()
Dim SQL As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    TieneEditorDeMenus = False
    SQL = "Select count(*) from usuarios.appmenus where aplicacion='Ariagro'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If miRsAux.Fields(0) > 0 Then TieneEditorDeMenus = True
        End If
    End If
    miRsAux.Close
        

ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub PoneMenusDelEditor()
Dim T As Control
Dim SQL As String
Dim C As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    
    SQL = "Select * from usuarios.appmenususuario where aplicacion='Ariagro' and codusu = " & Val(Right(CStr(vUsu.Codigo - vUsu.DevuelveAumentoPC), 3))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""

    While Not miRsAux.EOF
        If Not IsNull(miRsAux.Fields(3)) Then
            SQL = SQL & miRsAux.Fields(3) & "·"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
   
    If SQL <> "" Then
        SQL = "·" & SQL
        For Each T In Me.Controls
            If TypeOf T Is menu Then
                C = DevuelveCadenaMenu(T)
                C = "·" & C & "·"
                If InStr(1, SQL, C) > 0 Then T.visible = False
           
            End If
        Next
    End If
ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PoneBotonesdelToolbar()
    
    'clientes
    Me.Toolbar1.Buttons(1).visible = Me.Toolbar1.Buttons(1).visible And (mnComerGen.visible And mnComerAdm.visible And Me.mnComG_Admon(6).visible)
    Me.Toolbar1.Buttons(1).Enabled = Me.Toolbar1.Buttons(1).visible And (mnComerGen.Enabled And mnComerAdm.Enabled And Me.mnComG_Admon(6).Enabled)
    
    'variedades
    Me.Toolbar1.Buttons(2).visible = Me.Toolbar1.Buttons(2).visible And (Me.mnComerGen.visible And Me.mnComerPro.visible And Me.mnComG_Produc(4).visible)
    Me.Toolbar1.Buttons(2).Enabled = Me.Toolbar1.Buttons(2).visible And (Me.mnComerGen.Enabled And Me.mnComerPro.Enabled And Me.mnComG_Produc(4).Enabled)
    
    'forfaits
    Me.Toolbar1.Buttons(3).visible = Me.Toolbar1.Buttons(3).visible And (Me.mnComerGen.visible And Me.mnComerConfec.visible And Me.mnComG_Confec(8).visible)
    Me.Toolbar1.Buttons(3).Enabled = Me.Toolbar1.Buttons(3).visible And (Me.mnComerGen.Enabled And Me.mnComerConfec.Enabled And Me.mnComG_Confec(8).Enabled)
    
    'proveedores
    Me.Toolbar1.Buttons(4).visible = Me.Toolbar1.Buttons(4).visible And (Me.mnGesMater.visible And Me.mnComerGM_Gral.visible And Me.mnComGM_Gral(7).visible)
    Me.Toolbar1.Buttons(4).Enabled = Me.Toolbar1.Buttons(4).visible And (Me.mnGesMater.Enabled And Me.mnComerGM_Gral.Enabled And Me.mnComGM_Gral(7).Enabled)
    
    'articulos
    Me.Toolbar1.Buttons(5).visible = Me.Toolbar1.Buttons(5).visible And (Me.mnGesMater.visible And Me.mnComerGM_Gral.visible And Me.mnComGM_Gral(5).visible)
    Me.Toolbar1.Buttons(5).Enabled = Me.Toolbar1.Buttons(5).visible And (Me.mnGesMater.Enabled And Me.mnComerGM_Gral.Enabled And Me.mnComGM_Gral(5).Enabled)
    
    'pedidos de compra
    Me.Toolbar1.Buttons(7).visible = Me.Toolbar1.Buttons(7).visible And (Me.mnGesMater.visible And Me.mnComerGM_PedCom.visible And Me.mnComGM_PedCom(1).visible)
    Me.Toolbar1.Buttons(7).Enabled = Me.Toolbar1.Buttons(7).visible And (Me.mnGesMater.Enabled And Me.mnComerGM_PedCom.Enabled And Me.mnComGM_PedCom(1).Enabled)
    
    'albaranes de compra
    Me.Toolbar1.Buttons(8).visible = Me.Toolbar1.Buttons(8).visible And (Me.mnGesMater.visible And Me.mnComerGM_AlbCom.visible And Me.mnComGM_AlbCom(1).visible)
    Me.Toolbar1.Buttons(8).Enabled = Me.Toolbar1.Buttons(8).visible And (Me.mnGesMater.Enabled And Me.mnComerGM_AlbCom.Enabled And Me.mnComGM_AlbCom(1).Enabled)
    
    
    'facturas de compra
    Me.Toolbar1.Buttons(9).visible = Me.Toolbar1.Buttons(9).visible And (Me.mnGesMater.visible And Me.mnComerGM_AlbCom.visible And Me.mnComGM_AlbCom(6).visible)
    Me.Toolbar1.Buttons(9).Enabled = Me.Toolbar1.Buttons(9).visible And (Me.mnGesMater.Enabled And Me.mnComerGM_AlbCom.Enabled And Me.mnComGM_AlbCom(6).Enabled)
    
    'recepcion facturas de compra
    Me.Toolbar1.Buttons(10).visible = Me.Toolbar1.Buttons(10).visible And (Me.mnGesMater.visible And Me.mnComerGM_AlbCom.visible And Me.mnComGM_AlbCom(4).visible)
    Me.Toolbar1.Buttons(10).Enabled = Me.Toolbar1.Buttons(10).visible And (Me.mnGesMater.Enabled And Me.mnComerGM_AlbCom.Enabled And Me.mnComGM_AlbCom(4).Enabled)
    
    
    'palets
    Me.Toolbar1.Buttons(12).visible = Me.Toolbar1.Buttons(12).visible And (Me.mnGesVtas.visible And Me.mnComGV_PalPed(1).visible)
    Me.Toolbar1.Buttons(12).Enabled = Me.Toolbar1.Buttons(12).visible And (Me.mnGesVtas.Enabled And Me.mnComGV_PalPed(1).Enabled)
    
    'pedidos ventas
    Me.Toolbar1.Buttons(13).visible = Me.Toolbar1.Buttons(13).visible And (Me.mnGesVtas.visible And Me.mnComGV_PalPed(4).visible)
    Me.Toolbar1.Buttons(13).Enabled = Me.Toolbar1.Buttons(13).visible And (Me.mnGesVtas.Enabled And Me.mnComGV_PalPed(4).Enabled)
    
    'albaranes ventas
    Me.Toolbar1.Buttons(14).visible = Me.Toolbar1.Buttons(14).visible And (Me.mnGesVtas.visible And Me.mnComGV_PalPed(8).visible)
    Me.Toolbar1.Buttons(14).Enabled = Me.Toolbar1.Buttons(14).visible And (Me.mnGesVtas.Enabled And Me.mnComGV_PalPed(8).Enabled)
    
    'facturas ventas
    Me.Toolbar1.Buttons(17).visible = Me.Toolbar1.Buttons(17).visible And (Me.mnGesVtas.visible And Me.mnComerGV_Facturas.visible And Me.mnComGV_Facturas(1).visible)
    Me.Toolbar1.Buttons(17).Enabled = Me.Toolbar1.Buttons(17).visible And (Me.mnGesVtas.Enabled And Me.mnComerGV_Facturas.Enabled And Me.mnComGV_Facturas(1).Enabled)
    
    'facturas transporte
    Me.Toolbar1.Buttons(18).visible = Me.Toolbar1.Buttons(18).visible And (Me.mnGesVtas.visible And Me.mnComerGV_Portes.visible And Me.mnComGV_Portes(2).visible)
    Me.Toolbar1.Buttons(18).Enabled = Me.Toolbar1.Buttons(18).visible And (Me.mnGesVtas.Enabled And Me.mnComerGV_Portes.Enabled And Me.mnComGV_Portes(2).Enabled)
    
    'informes
    Me.Toolbar1.Buttons(19).visible = Me.Toolbar1.Buttons(19).visible And (Me.mnGesVtas.visible And Me.mnComerGV_Informes.visible)
    Me.Toolbar1.Buttons(19).Enabled = Me.Toolbar1.Buttons(19).visible And (Me.mnGesVtas.Enabled And Me.mnComerGV_Informes.Enabled)

    'cambio de campaña
    Me.Toolbar1.Buttons(21).visible = Me.Toolbar1.Buttons(21).visible And (Me.mnParametros(1).visible And Me.mnP_Generales(7).visible)
    Me.Toolbar1.Buttons(21).Enabled = Me.Toolbar1.Buttons(21).visible And (Me.mnParametros(1).Enabled And Me.mnP_Generales(7).Enabled)

End Sub


Private Function DevuelveCadenaMenu(ByRef T As Control) As String

On Error GoTo EDevuelveCadenaMenu
    DevuelveCadenaMenu = T.Name & "|"
    DevuelveCadenaMenu = DevuelveCadenaMenu & T.Index '& "|"   Monica:con esto no funcionaba
    Exit Function
EDevuelveCadenaMenu:
    Err.Clear
    
End Function

Private Sub LanzaHome(Opcion As String)
    Dim i As Integer
    Dim cad As String
    On Error GoTo ELanzaHome
    
    'Obtenemos la pagina web de los parametros
    CadenaDesdeOtroForm = DevuelveDesdeBD("websoporte", "sparam", "codparam", 1, "N")
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Falta configurar los datos en parametros.", vbExclamation
        Exit Sub
    End If
        
    i = FreeFile
    cad = ""
    Open App.path & "\lanzaexp.dat" For Input As #i
    Line Input #i, cad
    Close #i
    
    'Lanzamos
    If cad <> "" Then Shell cad & " " & CadenaDesdeOtroForm, vbMaximizedFocus
    
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, cad & vbCrLf & Err.Description
    CadenaDesdeOtroForm = ""
End Sub



Private Sub CargarImagen()

On Error GoTo eCargarImagen
    Me.Picture = LoadPicture(App.path & "\fondo.dat")
    Exit Sub
eCargarImagen:
    MuestraError Err.Number, "Error cargando imagen. LLame a soporte"
    End
End Sub

Private Sub PonerMenusNivelUsuario()
Dim b As Boolean

'    b = (vSesion.Codigo = 0)    'Sólo root y administrador
'
'    Me.mnE_Util(4).Enabled = b
'    Me.mnE_Util(4).visible = b
    
End Sub

Public Sub mnCambioEmpresa_Click()
    Dim AntUSU As Usuario

    If Not (Me.ActiveForm Is Nothing) Then
        MsgBox "Cierre todas las ventanas para poder cambiar de usuario", vbExclamation
        Exit Sub
    End If

    'Borramos temporal
    conn.Execute "Delete from zbloqueos where codusu = " & vUsu.Codigo


    CadenaDesdeOtroForm = vUsu.Login & "|" & vUsu.PasswdPROPIO & "|"
    
'    Set AntUSU = vUsu
'    Set vUsu = Nothing
    frmLogin.Show vbModal
'    If vUsu Is Nothing Then
'        Set vUsu = AntUSU
'        Set AntUSU = Nothing
'        Exit Sub
'    End If

    Screen.MousePointer = vbHourglass
    'Cerramos la conexion
    conn.Close
    If vParamAplic.NumeroConta <> 0 Then ConnConta.Close


    'Abre la conexión a BDatos:Ariges
    If AbrirConexion() = False Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
        End
    Else
        Set vParam = Nothing
        Set vParamAplic = Nothing
        'Carga Parametros Generales y Contables de la empresa
        LeerParametros
    End If


    'Abrir conexión a la BDatos de Contabilidad para acceder a
    'Tablas: Cuentas, Tipos IVA
    If vParamAplic.NumeroConta <> 0 Then
        If AbrirConexionConta() = False Then
            MsgBox "La aplicación no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
            End
        End If
    End If
    
    
    Set vEmpresa = Nothing
    'LeerEmpresaParametros
    
     'Carga los Datos Básicos de la empresa
    LeerDatosEmpresa
    
    
    'Carga los Niveles de cuentas de Contabilidad de la empresa
    If vParamAplic.NumeroConta <> 0 Then
        LeerNivelesEmpresa
    End If
    PonerDatosFormulario


    If vParamAplic.ContabilidadNueva And (vUsu.Nivel = 0 Or vUsu.Nivel = 1) Then FrasPendientesContabilizar False




    'Ponemos primera vez a false
    PrimeraVez = True
    Me.SetFocus

    Screen.MousePointer = vbDefault
End Sub



Private Sub PonerDatosVisiblesForm()
'Escribe texto de la barra de la aplicación
Dim cad As String
    cad = UCase(Mid(Format(Now, "dddd"), 1, 1)) & Mid(Format(Now, "dddd"), 2)
    cad = cad & ", " & Format(Now, "d")
    cad = cad & " de " & Format(Now, "mmmm")
    cad = cad & " de " & Format(Now, "yyyy")
    cad = "    " & cad & "    "
    Me.StatusBar1.Panels(7).Text = cad
    
    '
    Me.StatusBar1.Panels(2).Text = vUsu.CadenaConexion
    
    '[Monica]26/05/2014: solo en el caso de natural y en el de img quitamos lo de campaña anterior
    If vParamAplic.Cooperativa = 15 Or vParamAplic.Cooperativa = 9 Then
        Me.StatusBar1.Panels(4).visible = False
    Else
        If Not EsCampanyaActual(vEmpresa.BDAriagro) Then
            Me.StatusBar1.Panels(4).visible = True
        Else
            Me.StatusBar1.Panels(4).visible = False
        End If
    End If
    
    cad = ""
    If vParam Is Nothing Then
        Caption = "AriAgro - Comercial " & " v." & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & " FALTA CONFIGURAR"
    Else
        Caption = "AriAgro - Comercial " & " v." & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & vParam.NombreEmpresa & cad & _
                  "   -   " & vEmpresa.nomresum & "   -   Fechas: " & vParam.FecIniCam & " - " & vParam.FecFinCam & "   -   Usuario: " & vUsu.Nombre
    End If
End Sub




