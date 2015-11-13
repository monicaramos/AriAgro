VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmANECOOPExped 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Expedientes de ANECOOP"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   12780
   Icon            =   "frmANECOOPExped.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmANECOOPExped.frx":000C
   ScaleHeight     =   8670
   ScaleWidth      =   12780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAux1 
      Caption         =   "Pagos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3330
      Left            =   120
      TabIndex        =   41
      Top             =   4530
      Width           =   12435
      Begin VB.CheckBox chkAux 
         BackColor       =   &H80000005&
         Height          =   255
         Index           =   0
         Left            =   10140
         TabIndex        =   85
         Tag             =   "IdContab|N|N|0|1|anecoop_pago|idcontab|||"
         Top             =   2250
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   3000
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   78
         Text            =   "Tip"
         Top             =   2250
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   2
         Left            =   9780
         MaskColor       =   &H00000000&
         TabIndex        =   77
         ToolTipText     =   "Buscar fecha"
         Top             =   2250
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   1
         Left            =   8790
         MaskColor       =   &H00000000&
         TabIndex        =   76
         ToolTipText     =   "Buscar fecha"
         Top             =   2250
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   0
         Left            =   6060
         MaskColor       =   &H00000000&
         TabIndex        =   75
         ToolTipText     =   "Buscar fecha"
         Top             =   2250
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   4905
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   49
         Tag             =   "Fecha Factura|F|S|||anecoop_pago|fecha_factura|dd/mm/yyyy||"
         Text            =   "Fecha Fra"
         Top             =   2250
         Width           =   1200
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   990
         MaxLength       =   18
         TabIndex        =   42
         Tag             =   "Expediente|T|S|||anecoop_pago|expediente_id||S|"
         Text            =   "Expediente"
         Top             =   2250
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   3645
         MaxLength       =   10
         TabIndex        =   47
         Tag             =   "Nro Factura|T|S|||anecoop_pago|num_factura||S|"
         Text            =   "factura"
         Top             =   2250
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   6255
         MaxLength       =   8
         TabIndex        =   50
         Tag             =   "Nro Liquidacion|N|S|||anecoop_pago|num_liquidacion|#######0||"
         Text            =   "Nro Liquidac"
         Top             =   2250
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   1740
         MaxLength       =   18
         TabIndex        =   43
         Tag             =   "Expediente|T|S|||anecoop_pago|expediente_pagoid||S|"
         Text            =   "Linea"
         Top             =   2250
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   7155
         MaxLength       =   12
         TabIndex        =   52
         Tag             =   "Importe|N|N|||anecoop_pago|importe|###,###,##0.00||"
         Text            =   "Importe"
         Top             =   2250
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   8010
         MaxLength       =   10
         TabIndex        =   54
         Tag             =   "Fecha Pago|F|S|||anecoop_pago|fecha_pago|dd/mm/yyyy||"
         Text            =   "Fecha Pago"
         Top             =   2250
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   8
         Left            =   9000
         MaxLength       =   10
         TabIndex        =   56
         Tag             =   "Fecha Pago Seccion|F|S|||anecoop_pago|fecha_pago_sc|dd/mm/yyyy||"
         Text            =   "Fec Pago S"
         Top             =   2250
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   2430
         MaxLength       =   2
         TabIndex        =   45
         Tag             =   "Tipo Pago|T|S|||anecoop_pago|tipo_pago|||"
         Text            =   "Ti"
         Top             =   2250
         Visible         =   0   'False
         Width           =   420
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   1
         Left            =   240
         TabIndex        =   44
         Top             =   315
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmANECOOPExped.frx":0A0E
         Height          =   2130
         Left            =   240
         TabIndex        =   46
         Top             =   810
         Width           =   11310
         _ExtentX        =   19950
         _ExtentY        =   3757
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adoaux 
         Height          =   330
         Index           =   1
         Left            =   1455
         Top             =   315
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3705
      Left            =   120
      TabIndex        =   35
      Top             =   690
      Width           =   12435
      Begin VB.CheckBox Check1 
         Caption         =   "Bloqueado"
         Height          =   255
         Index           =   0
         Left            =   9150
         TabIndex        =   86
         Tag             =   "Bloqueado|N|N|0|1|anecoop|bloqueado|||"
         Top             =   3210
         Width           =   1065
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D6DEFE&
         Height          =   315
         Index           =   22
         Left            =   11280
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   84
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D6DEFE&
         Height          =   315
         Index           =   23
         Left            =   10260
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   83
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00D6DEFE&
         Height          =   315
         Index           =   13
         Left            =   7470
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   82
         Top             =   2190
         Width           =   1665
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00D6DEFE&
         Height          =   315
         Index           =   14
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   81
         Top             =   2190
         Width           =   2385
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00D6DEFE&
         Height          =   315
         Index           =   16
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   80
         Top             =   2190
         Width           =   2385
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00D6DEFE&
         Height          =   315
         Index           =   17
         Left            =   210
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   79
         Top             =   2190
         Width           =   2385
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   26
         Left            =   9120
         MaxLength       =   6
         TabIndex        =   13
         Tag             =   "% Iva Liquidaciónl|N|S|||anecoop|porcent_iva_liq|##0.00||"
         Top             =   990
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   24
         Left            =   9240
         MaxLength       =   15
         TabIndex        =   25
         Tag             =   "Precio Comercial|N|S|||anecoop|precio_comercial|#,###,##0.000||"
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   23
         Left            =   10260
         MaxLength       =   15
         TabIndex        =   26
         Tag             =   "Peso Neto|N|S|||anecoop|peso_neto|###,###,##0||"
         Text            =   "000000000"
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   22
         Left            =   11280
         MaxLength       =   10
         TabIndex        =   27
         Tag             =   "Cajas|N|S|||anecoop|ncajas|###,###,##0||"
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   21
         Left            =   2640
         MaxLength       =   25
         TabIndex        =   23
         Tag             =   "Tipo Palet|T|S|||anecoop|nombre_tipo_palet|||"
         Text            =   "Text1"
         Top             =   2760
         Width           =   2385
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   20
         Left            =   5040
         MaxLength       =   50
         TabIndex        =   24
         Tag             =   "Artículo|T|S|||anecoop|nombre_articulo|||"
         Text            =   "Text1"
         Top             =   2760
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   19
         Left            =   2850
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "Codigo Campaña|T|S|||anecoop|codigo_campanya||S|"
         Text            =   "Text1 7"
         Top             =   390
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   18
         Left            =   2010
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "Linea|T|S|||anecoop|linea_expediente||S|"
         Text            =   "Text1 7"
         Top             =   390
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   17
         Left            =   210
         MaxLength       =   25
         TabIndex        =   16
         Tag             =   "Variedad|T|S|||anecoop|nombre_variedad|||"
         Text            =   "Text1"
         Top             =   1830
         Width           =   2385
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   16
         Left            =   2640
         MaxLength       =   25
         TabIndex        =   17
         Tag             =   "Marca|T|S|||anecoop|nombre_marca|||"
         Text            =   "Text1"
         Top             =   1830
         Width           =   2385
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   15
         Left            =   210
         MaxLength       =   25
         TabIndex        =   22
         Tag             =   "Material|T|S|||anecoop|nombre_material|||"
         Text            =   "Text1"
         Top             =   2760
         Width           =   2385
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   14
         Left            =   5040
         MaxLength       =   18
         TabIndex        =   18
         Tag             =   "Confeccion|T|S|||anecoop|nombre_confeccion|||"
         Text            =   "0000000000000000000000000"
         Top             =   1830
         Width           =   2385
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   13
         Left            =   7470
         MaxLength       =   25
         TabIndex        =   19
         Tag             =   "Categoria|T|S|||anecoop|nombre_categoria|||"
         Text            =   "Text1"
         Top             =   1830
         Width           =   1665
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   12
         Left            =   9810
         MaxLength       =   15
         TabIndex        =   14
         Tag             =   "Importe Iva Liquidaciónl|N|S|||anecoop|importe_iva_liq|#,###,##0.00||"
         Top             =   990
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   11
         Left            =   5580
         MaxLength       =   9
         TabIndex        =   10
         Tag             =   "Factura|T|S|||anecoop|fra_liq|||"
         Top             =   1020
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   10
         Left            =   10830
         MaxLength       =   15
         TabIndex        =   21
         Tag             =   "Valor Confeccion|N|S|||anecoop|valor_confeccion|#,###,##0.00||"
         Top             =   1830
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   9240
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "Valor Mercancia|N|S|||anecoop|valor_mercancia|#,###,##0.00||"
         Top             =   1830
         Width           =   1515
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   10890
         MaxLength       =   15
         TabIndex        =   15
         Tag             =   "Importe Liquidaciónl|N|S|||anecoop|importe_liq|#,###,##0.00||"
         Top             =   990
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   7980
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "Fecha Liquidacion|F|S|||anecoop|fecha_liq|dd/mm/yyyy||"
         Top             =   990
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   6810
         MaxLength       =   5
         TabIndex        =   11
         Tag             =   "Liquidacion|N|S|||anecoop|num_liq|||"
         Top             =   990
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   10890
         MaxLength       =   12
         TabIndex        =   8
         Tag             =   "Matricula|T|S|||anecoop|matricula|||"
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   25
         Left            =   6810
         MaxLength       =   7
         TabIndex        =   5
         Tag             =   "Linea Albarán|N|S|||anecoop|numlinea|||"
         Text            =   "Text1 7"
         Top             =   420
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   9810
         MaxLength       =   7
         TabIndex        =   7
         Tag             =   "Numero de Pedido|N|S|||anecoop|n_pedido_aneccop|######0||"
         Text            =   "Text1 7"
         Top             =   420
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   5610
         MaxLength       =   12
         TabIndex        =   4
         Tag             =   "Nro Salida|T|S|||anecoop|numero_salida_cooperativa|||"
         Text            =   "Text1 7"
         Top             =   420
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   3660
         MaxLength       =   5
         TabIndex        =   3
         Tag             =   "Periodo|T|S|||anecoop|periodo|||"
         Text            =   "Text1 7"
         Top             =   390
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   7965
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Fecha Salida|F|S|||anecoop|fecha_salida|dd/mm/yyyy||"
         Top             =   420
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   27
         Left            =   840
         MaxLength       =   50
         TabIndex        =   9
         Tag             =   "Cliente|T|S|||anecoop|nombre_cliente|||"
         Text            =   "00000000000000000000000000000000000000000000000000"
         Top             =   1020
         Width           =   4635
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   180
         MaxLength       =   18
         TabIndex        =   0
         Tag             =   "Expediente|T|S|||anecoop|expediente_id||S|"
         Text            =   "Text1 7"
         Top             =   390
         Width           =   1725
      End
      Begin VB.Label Label1 
         Caption         =   "%Iva"
         Height          =   255
         Index           =   15
         Left            =   9180
         TabIndex        =   74
         Top             =   750
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Campaña"
         Height          =   255
         Index           =   14
         Left            =   2850
         TabIndex        =   73
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Linea"
         Height          =   255
         Index           =   13
         Left            =   2010
         TabIndex        =   72
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Iva Liquid."
         Height          =   255
         Index           =   12
         Left            =   9840
         TabIndex        =   71
         Top             =   750
         Width           =   1020
      End
      Begin VB.Label Label29 
         Caption         =   "Precio"
         Height          =   255
         Index           =   8
         Left            =   9240
         TabIndex        =   70
         Top             =   2520
         Width           =   945
      End
      Begin VB.Label Label29 
         Caption         =   "Peso Neto"
         Height          =   255
         Index           =   7
         Left            =   10320
         TabIndex        =   69
         Top             =   2520
         Width           =   915
      End
      Begin VB.Label Label29 
         Caption         =   "Cajas"
         Height          =   255
         Index           =   6
         Left            =   11310
         TabIndex        =   68
         Top             =   2490
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "Tipo Palet"
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   67
         Top             =   2550
         Width           =   1125
      End
      Begin VB.Label Label29 
         Caption         =   "Artículo"
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   66
         Top             =   2490
         Width           =   1125
      End
      Begin VB.Label Label29 
         Caption         =   "Variedad"
         Height          =   255
         Index           =   3
         Left            =   210
         TabIndex        =   65
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label29 
         Caption         =   "Marca"
         Height          =   255
         Index           =   2
         Left            =   2670
         TabIndex        =   64
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label29 
         Caption         =   "Material"
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   63
         Top             =   2520
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Factura "
         Height          =   255
         Index           =   11
         Left            =   5580
         TabIndex        =   62
         Top             =   780
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Valor Confección"
         Height          =   255
         Index           =   10
         Left            =   10890
         TabIndex        =   61
         Top             =   1560
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "Valor Mercancia"
         Height          =   255
         Index           =   9
         Left            =   9270
         TabIndex        =   60
         Top             =   1560
         Width           =   1170
      End
      Begin VB.Image imgFec 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   8820
         Picture         =   "frmANECOOPExped.frx":0A23
         ToolTipText     =   "Buscar fecha"
         Top             =   750
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Liq."
         Height          =   255
         Index           =   7
         Left            =   7980
         TabIndex        =   59
         Top             =   750
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Liquidación"
         Height          =   255
         Index           =   6
         Left            =   6840
         TabIndex        =   58
         Top             =   750
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Matrícula"
         Height          =   255
         Index           =   5
         Left            =   10965
         TabIndex        =   57
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Linea Albarán"
         Height          =   255
         Index           =   4
         Left            =   6810
         TabIndex        =   55
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Pedido"
         Height          =   255
         Index           =   3
         Left            =   9810
         TabIndex        =   53
         Top             =   180
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Albarán"
         Height          =   255
         Index           =   2
         Left            =   5610
         TabIndex        =   51
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Liquidacion"
         Height          =   255
         Index           =   17
         Left            =   10920
         TabIndex        =   48
         Top             =   750
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Categoría"
         Height          =   255
         Index           =   8
         Left            =   7500
         TabIndex        =   40
         Top             =   1590
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   29
         Left            =   7980
         TabIndex        =   39
         Top             =   180
         Width           =   705
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   8820
         Picture         =   "frmANECOOPExped.frx":0AAE
         ToolTipText     =   "Buscar fecha"
         Top             =   150
         Width           =   240
      End
      Begin VB.Label Label29 
         Caption         =   "Confección"
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   38
         Top             =   1590
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   37
         Top             =   1065
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Expediente"
         Height          =   255
         Index           =   28
         Left            =   210
         TabIndex        =   36
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   31
      Top             =   8010
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   32
         Top             =   180
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11535
      TabIndex        =   29
      Top             =   8100
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10320
      TabIndex        =   28
      Top             =   8100
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   12780
      _ExtentX        =   22543
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Añadir"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Desdoblar Expediente"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   10680
         TabIndex        =   34
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11520
      TabIndex        =   30
      Top             =   8100
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   3000
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   240
      Top             =   8040
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   240
      Top             =   8070
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnDesdoblar 
         Caption         =   "Desdoblar Expediente"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmANECOOPExped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCC As frmCal
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmAux As frmANECOOPAux
Attribute frmAux.VB_VarHelpID = -1

Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean


Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal
Dim indice As Byte

Dim TipoFactura As Byte
Private BuscaChekc As String

Private Sub btnBuscar_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object
    
    TerminaBloquear
    
    Select Case Index
        Case 0, 1, 2 'fecha factura, pago y pago seccion
        
            Select Case Index
                Case 0
                    indice = 4
                Case 1
                    indice = 7
                Case 2
                    indice = 8
            End Select
            
            
            Set frmCC = New frmCal
            
            esq = btnBuscar(Index).Left
            dalt = btnBuscar(Index).Top
                
            Set obj = btnBuscar(Index).Container
              
              While btnBuscar(Index).Parent.Name <> obj.Name
                    esq = esq + obj.Left
                    dalt = dalt + obj.Top
                    Set obj = obj.Container
              Wend
            
            menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
        
            frmCC.Left = esq + btnBuscar(Index).Parent.Left + 30
            frmCC.Top = dalt + btnBuscar(Index).Parent.Top + btnBuscar(Index).Height + menu - 40
        
            btnBuscar(0).Tag = indice '<===
            ' *** repasar si el camp es txtAux o Text1 ***
            If txtAux(indice).Text <> "" Then frmCC.NovaData = txtAux(indice).Text
        
            frmCC.Show vbModal
            Set frmCC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonerFoco txtAux(CByte(btnBuscar(0).Tag) + 1) '<===
            ' ********************************************
        
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1

End Sub


Private Sub check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim I As Integer

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda

        Case 3  'AÑADIR
            If DatosOk Then InsertarCabecera

        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaCabecera Then
                    espera 0.2
                    TerminaBloquear
                    PosicionarData
                    PonerCampos
                    PonerCamposLineas
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            Select Case ModificaLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    If ModificarLinea Then PosicionarData
            End Select
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 0, 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            LLamaLineas Modo, 0, "DataGrid2"
            PonerFoco Text1(3)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(3)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid3.AllowAddNew = False
                If Not Adoaux(1).Recordset.EOF Then Adoaux(1).Recordset.MoveFirst
            End If
            ModificaLineas = 0
            LLamaLineas Modo, 0, "DataGrid3"
            PonerModo 2
            DataGrid3.Enabled = True
            If Not Data1.Recordset.EOF Then _
                Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
            'Habilitar las opciones correctas del menu segun Modo
            PonerModoOpcionesMenu (Modo)
            PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
            DataGrid3.Enabled = True
            PonerFocoGrid DataGrid3
    
    End Select
End Sub

Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    
    
    PonerModo 3
    
    LimpiarDataGrids
    
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions
End Sub


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select anecoop.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla & " order by expediente_id "
'        CadenaConsulta = CadenaConsulta & " WHERE scafac.codtipom='" & CodTipoMov & "'"
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim DeVarios As Boolean

    
    PonerModo 4
    
    PonerFoco Text1(2) '*** 1r camp visible que siga PK ***
        
End Sub


Private Sub BotonModificarLinea(Index As Integer)
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte

    On Error GoTo eModificarLinea


    ModificaLineas = 2 'Modificar

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then  '1= Insertar
        TerminaBloquear
        Exit Sub
    End If
    
    
    PonerModo 5, Index
 
    
    vWhere = "expediente_id = " & DBSet(Adoaux(1).Recordset!expediente_id, "T")
    vWhere = vWhere & " and expediente_pagoid=" & DBSet(Adoaux(1).Recordset!expediente_pagoid, "T")
    If Not BloqueaRegistro("anecoop_pago", vWhere) Then
        TerminaBloquear
        Exit Sub
    End If
    If DataGrid3.Bookmark < DataGrid3.FirstRow Or DataGrid3.Bookmark > (DataGrid3.FirstRow + DataGrid3.VisibleRows - 1) Then
        J = DataGrid3.Bookmark - DataGrid3.FirstRow
        DataGrid3.Scroll 0, J
        DataGrid3.Refresh
    End If
    
'    anc = ObtenerAlto(Me.DataGrid1)
    anc = DataGrid3.Top
    If DataGrid3.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 10
    End If

    txtAux(0).Text = DataGrid3.Columns(0).Text ' albaran
    txtAux(1).Text = DataGrid3.Columns(1).Text ' linea
    txtAux(2).Text = DataGrid3.Columns(2).Text ' almacen
    txtAux2(2).Text = DataGrid3.Columns(3).Text
    txtAux(3).Text = DataGrid3.Columns(4).Text ' articulo
    txtAux(4).Text = DataGrid3.Columns(5).Text ' cantidad
    txtAux(5).Text = DataGrid3.Columns(6).Text ' precio
    txtAux(6).Text = DataGrid3.Columns(7).Text ' dtolinea
    txtAux(7).Text = DataGrid3.Columns(8).Text ' importe
    txtAux(8).Text = DataGrid3.Columns(9).Text ' codigo de iva
    

    LLamaLineas ModificaLineas, anc, "DataGrid3"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid3.Enabled = True
    
    PonerFoco txtAux(2)
    Me.DataGrid3.Enabled = False


eModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim b As Boolean
    
    Select Case grid
        Case "DataGrid3"
            DeseleccionaGrid Me.DataGrid3
            b = (xModo = 1 Or xModo = 2)
            For jj = 2 To 8
                txtAux(jj).Height = DataGrid3.RowHeight - 10
                txtAux(jj).Top = alto + 5
                txtAux(jj).visible = b
                txtAux(jj).Enabled = b
            Next jj
            txtAux2(2).Height = DataGrid3.RowHeight - 10
            txtAux2(2).Top = alto + 5
            txtAux2(2).visible = b
            txtAux2(2).Enabled = False
'            txtAux2(2).BackColor = &H80000018
            For jj = 0 To 2
                btnBuscar(jj).Height = DataGrid3.RowHeight - 10
                btnBuscar(jj).Top = alto + 5
                btnBuscar(jj).visible = b
                btnBuscar(jj).Enabled = b
            Next jj
    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim cad As String
'Dim vTipoMov As CTiposMov

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then Exit Sub
    
    cad = "Cabecera de Expedientes." & vbCrLf
    cad = cad & "-------------------------------------      " & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el Expediente:            "
    cad = cad & vbCrLf & "Nº Expediente:  " & Text1(0).Text
    cad = cad & vbCrLf & "Linea:  " & Text1(18).Text
    cad = cad & vbCrLf & "Campaña:  " & Text1(19).Text

    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If
        
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminar Expediente", Err.Description
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        DataGrid3.Enabled = True
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
     
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    

    ' ICONITOS DE LA BARRA
    btnPrimero = 14
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(4).Image = 3   'Insertar
        .Buttons(5).Image = 4   'Modificar
        .Buttons(6).Image = 5   'Borrar
        .Buttons(8).Image = 19  'Desdoblar Expediente
        .Buttons(9).Image = 10  'Impresión de albaran
        .Buttons(11).Image = 11  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For kCampo = 1 To 1
        With Me.ToolAux(kCampo)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next kCampo
   ' ***********************************
    
    LimpiarCampos   'Limpia los campos TextBox

    VieneDeBuscar = False
    
        
    '## A mano
    NombreTabla = "anecoop"
    NomTablaLineas = "anecoop_pago" 'Tabla lineas
    Ordenacion = " ORDER BY anecoop.expediente_id, anecoop.linea_expediente, anecoop.codigo_campanya"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CadenaConsulta = "select * from anecoop "
    CadenaConsulta = CadenaConsulta & " where anecoop.expediente_id = -1"
    
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
'        CargaGrid DataGrid1, Data2, False
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
    
    PrimeraVez = False
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    Me.Check1(0).Value = 0
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        
        
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
     Text1(indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
     txtAux(indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub


Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. de Cliente
'            indice = 3
'            PonerFoco Text1(indice)
'            Set frmCli = New frmClientes
'            frmCli.DatosADevolverBusqueda = "0|1|2|"
'            frmCli.Show vbModal
'            Set frmCli = Nothing
'            PonerFoco Text1(indice)
        
        Case 1 'Forma de Pago
'            indice = 4
'            PonerFoco Text1(indice)
'            Set frmFPag = New frmManFpago
'            frmFPag.DatosADevolverBusqueda = "0|1|"
'            frmFPag.Show vbModal
'            Set frmFPag = Nothing
'            PonerFoco Text1(indice)
            
    End Select
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmC = New frmCal
    
    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top
        
    Set obj = imgFec(Index).Container
      
      While imgFec(Index).Parent.Name <> obj.Name
            esq = esq + obj.Left
            dalt = dalt + obj.Top
            Set obj = obj.Container
      Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40


    If Index = 0 Then
        indice = 4
    Else
        indice = 7
    End If

    If Text1(indice).Text <> "" Then frmC.NovaData = Text1(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
     PonerFoco Text1(indice) '<===
    ' ********************************************
End Sub




Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnDesdoblar_Click()

    If ComprobarCero(Text1(3).Text) = 0 Then Exit Sub

    BotonDesdoblar
End Sub

Private Sub BotonDesdoblar()
Dim Sql As String


    Sql = "delete from anecoopaux where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    Sql = "insert into anecoopaux (codusu,numalbar,numlinea,numcajas,pesoneto) "
    Sql = Sql & " select " & vUsu.Codigo & ", numalbar, numlinea, numcajas, pesoneto from albaran_variedad where numalbar = " & DBSet(Text1(3).Text, "N")
    Sql = Sql & " and (expediente is null or expediente = '') "
    
    conn.Execute Sql
    
    Sql = "select count(*) from anecoopaux where codusu = " & vUsu.Codigo
    If TotalRegistros(Sql) <> 0 Then
    
        Set frmAux = New frmANECOOPAux
        frmANECOOPAux.Show vbModal
        Set frmAux = Nothing

        If MsgBox("¿ Desea continuar con el proceso de generar expedientes ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
        
        If CrearExpedientes Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
        End If
    Else
        MsgBox "No hay lineas de este albarán pendientes de asignar", vbExclamation

    End If

End Sub

Private Function CrearExpedientes() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Numexped As String
Dim I As Integer
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim RS3 As ADODB.Recordset
Dim PNetoTot As Currency
Dim BaseTot As Currency
Dim IvaTot As Currency
Dim ValorTot As Currency

Dim vPNeto As Currency
Dim vBase As Currency
Dim vIva As Currency
Dim vValor As Currency

Dim vBase1 As Currency
Dim vIva1 As Currency
Dim vValor1 As Currency

Dim vImporte As Currency
Dim vImporte1 As Currency
Dim ImporteTot As Currency
Dim PNeto As Currency
    
    
    On Error GoTo eCrearExpedientes

    CrearExpedientes = False

    conn.BeginTrans

    Sql = "select * from anecoopaux where codusu = " & vUsu.Codigo
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Numexped = Text1(0).Text
    
    If Not RS.EOF Then
        BaseTot = 0
        IvaTot = 0
        ValorTot = 0
    
        Sql2 = "select peso_neto,importe_liq,importe_iva_liq,valor_mercancia from anecoop where expediente_id = " & DBSet(Text1(0).Text, "T")
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        vPNeto = 0
        vBase = 0
        vIva = 0
        vValor = 0
        If Not Rs2.EOF Then
            vPNeto = DevuelveValor("select sum(pesoneto) from anecoopaux where codusu = " & vUsu.Codigo) 'DBLet(Rs2!peso_neto, "N")
            vBase = DBLet(Rs2!importe_liq, "N")
            vIva = DBLet(Rs2!importe_iva_liq, "N")
            vValor = DBLet(Rs2!valor_mercancia, "N")
        End If
        Set Rs2 = Nothing
        
        vBase1 = 0
        vIva1 = 0
        vValor1 = 0
        If vPNeto <> 0 Then
            vBase1 = Round2(vBase * DBLet(RS!Pesoneto, "N") / vPNeto, 2)
            vIva1 = Round2(vIva * DBLet(RS!Pesoneto, "N") / vPNeto, 2)
            vValor1 = Round2(vValor * DBLet(RS!Pesoneto, "N") / vPNeto, 2)
            
            BaseTot = BaseTot + vBase1
            IvaTot = IvaTot + vIva1
            ValorTot = ValorTot + vValor1
        End If
    
        Sql = "update anecoop set ncajas =" & DBSet(RS!NumCajas, "N")
        Sql = Sql & ", peso_neto = " & DBSet(RS!Pesoneto, "N")
        Sql = Sql & ", numlinea = " & DBSet(RS!numlinea, "N")
        Sql = Sql & ", importe_liq = " & DBSet(vBase1, "N")
        Sql = Sql & ", importe_iva_liq = " & DBSet(vIva1, "N")
        Sql = Sql & ", valor_mercancia = " & DBSet(vValor1, "N")
        Sql = Sql & ", bloqueado = 1 "
        Sql = Sql & " where expediente_id = " & DBSet(Text1(0).Text, "T")
        Sql = Sql & " and linea_expediente = " & DBSet(Text1(18).Text, "T")
        Sql = Sql & " and codigo_campanya = " & DBSet(Text1(19).Text, "T")
        
        conn.Execute Sql
        
        RS.MoveNext
    
        ' resto de expedientes, los creados
        I = 0
        While Not RS.EOF
            I = I + 1
            ' concateno 18 ceros al inicio y cojo los 18 de la derecha (esto es por si el expediente no está relleno a ceros por la izquierda)
            Numexped = Right("000000000000000000" & Text1(0).Text, 18)
            ' le quito la primera posicion para poner un 1, 2 o el que sea
            Numexped = I & Mid(Numexped, Len(CStr(I)) + 1, 18 - Len(CStr(I)))
            
            
            vBase1 = 0
            vIva1 = 0
            vValor1 = 0
            If vPNeto <> 0 Then
                vBase1 = Round2(vBase * DBLet(RS!Pesoneto, "N") / vPNeto, 2)
                vIva1 = Round2(vIva * DBLet(RS!Pesoneto, "N") / vPNeto, 2)
                vValor1 = Round2(vValor * DBLet(RS!Pesoneto, "N") / vPNeto, 2)
                
                BaseTot = BaseTot + vBase1
                IvaTot = IvaTot + vIva1
                ValorTot = ValorTot + vValor1
            End If
            
            ' ahora insertamos
            Sql = "insert into anecoop (expediente_id,pdls_id,linea_expediente,codigo_campanya,periodo,codigo_cooperativa_gestora,nombre_gestora,codigo_cooperativa_carga,nombre_carga,pto_carga,"
            Sql = Sql & "nombre_pto_carga,numero_salida_cooperativa,nlinea_salida_cooperativa,n_pedido_aneccop,n_pedido,n_linea,tipo_expediente,estado_coop_expediente,expediente_reenviado,"
            Sql = Sql & "codigo_delegacion,nombre_delegacion,codigo_area,nombre_area,fecha_salida,matricula,tipo_vehiculo,categoria,nombre_categoria,codigo_confeccion,nombre_confeccion,"
            Sql = Sql & "codigo_uventa,nombre_uventa,codigo_etrasnsporte,nombre_etransporte,codigo_modelo,nombre_modelo,codigo_material,nombre_material,codigo_marca,nombre_marca,codigo_variedad,"
            Sql = Sql & "nombre_variedad,codigo_articulo,nombre_articulo,ean13,codigo_tipo_articulo,codigo_cliente,nombre_cliente,codigo_pais,codigo_divisa_cliente,codigo_producto,"
            Sql = Sql & "nombre_producto,codigo_calibre,nombre_calibre,npalet,tipo_palet,nombre_tipo_palet,ncajas,peso_neto,unidad_peso,precio_comercial,codigo_divisa,liquidado_son,"
            Sql = Sql & "estado_coop_liquidacion,liquidacion_agrupada_son,codigo_grupo_liquidacion,num_liq,fra_liq,fra_liq_complementa,carta_liq,registro_liq,importe_liq,"
            Sql = Sql & "importe_iva_liq,tipo_iva_liq,fecha_liq,fecha_cambio_liq,cambio_liq,fecha_sc_liq,porcent_iva_liq,importe2_iva_liq,cobrado_son,pagado_son,valor_mercancia,valor_confeccion,"
            Sql = Sql & "numero_salida_anecoop,salidalineaid,numlinea,bloqueado) "
            Sql = Sql & " select " & DBSet(Numexped, "T") & ",pdls_id,linea_expediente,codigo_campanya,periodo,codigo_cooperativa_gestora,nombre_gestora,codigo_cooperativa_carga,nombre_carga,pto_carga,"
            Sql = Sql & "nombre_pto_carga,numero_salida_cooperativa,nlinea_salida_cooperativa,n_pedido_aneccop,n_pedido,n_linea,tipo_expediente,estado_coop_expediente,expediente_reenviado,"
            Sql = Sql & "codigo_delegacion,nombre_delegacion,codigo_area,nombre_area,fecha_salida,matricula,tipo_vehiculo,categoria,nombre_categoria,codigo_confeccion,nombre_confeccion,"
            Sql = Sql & "codigo_uventa,nombre_uventa,codigo_etrasnsporte,nombre_etransporte,codigo_modelo,nombre_modelo,codigo_material,nombre_material,codigo_marca,nombre_marca,codigo_variedad,"
            Sql = Sql & "nombre_variedad,codigo_articulo,nombre_articulo,ean13,codigo_tipo_articulo,codigo_cliente,nombre_cliente,codigo_pais,codigo_divisa_cliente,codigo_producto,"
            Sql = Sql & "nombre_producto,codigo_calibre,nombre_calibre,npalet,tipo_palet,nombre_tipo_palet," & DBSet(RS!NumCajas, "N") & "," & DBSet(RS!Pesoneto, "N") & ",unidad_peso,precio_comercial,codigo_divisa,liquidado_son,"
            Sql = Sql & "estado_coop_liquidacion,liquidacion_agrupada_son,codigo_grupo_liquidacion,num_liq,fra_liq,fra_liq_complementa,carta_liq,registro_liq," & DBSet(vBase1, "N") & ","
            Sql = Sql & DBSet(vIva1, "N") & ",tipo_iva_liq,fecha_liq,fecha_cambio_liq,cambio_liq,fecha_sc_liq,porcent_iva_liq,importe2_iva_liq,cobrado_son,pagado_son," & DBSet(vValor1, "N") & ",valor_confeccion,"
            Sql = Sql & "numero_salida_anecoop,salidalineaid," & DBSet(RS!numlinea, "N") & ",bloqueado "
            Sql = Sql & " from anecoop "
            Sql = Sql & " where expediente_id = " & DBSet(Text1(0).Text, "T")
            Sql = Sql & " and linea_expediente = " & DBSet(Text1(18).Text, "T")
            Sql = Sql & " and codigo_campanya = " & DBSet(Text1(19).Text, "T")
            
            conn.Execute Sql
            
            
            
            RS.MoveNext
        Wend
        Set RS = Nothing
    
        If (BaseTot <> vBase Or IvaTot <> vIva Or ValorTot <> vValor) Then
            ' los redondeos los hacemos en el primer registro
            Sql = "update anecoop set "
            Sql = Sql & " importe_liq = importe_liq + " & DBSet(vBase - BaseTot, "N")
            Sql = Sql & ", importe_iva_liq = importe_iva_liq + " & DBSet(vIva - IvaTot, "N")
            Sql = Sql & ", valor_mercancia = valor_mercancia + " & DBSet(vValor - ValorTot, "N")
            Sql = Sql & " where expediente_id = " & DBSet(Text1(0).Text, "T")
            Sql = Sql & " and linea_expediente = " & DBSet(Text1(18).Text, "T")
            Sql = Sql & " and codigo_campanya = " & DBSet(Text1(19).Text, "T")
            
            conn.Execute Sql
        End If
        
        'actualizamos los pagos si los hay
        Sql2 = "select * from anecoop_pago where expediente_id = " & DBSet(Text1(0).Text, "T")
        Sql2 = Sql2 & " order by expediente_pagoid "
        
        If TotalRegistrosConsulta(Sql2) <> 0 Then
            ' para cada linea de pago prorrateamos
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs2.EOF
            
                PNeto = DevuelveValor("select peso_neto from anecoop where expediente_id = " & DBSet(Text1(0).Text, "T"))
            
                vImporte = DBLet(Rs2!Importe, "N")
                
                vImporte1 = Round2(PNeto * vImporte / vPNeto, 2)
                
                ImporteTot = vImporte1
                
                
                'la primera linea la actualizamos
                Sql = "update anecoop_pago set importe = " & DBSet(vImporte1, "N")
                Sql = Sql & " where expediente_id = " & DBSet(Text1(0).Text, "T")
                Sql = Sql & " and expediente_pagoid = " & DBSet(Rs2!expediente_pagoid, "T")
                
                conn.Execute Sql
            
            
                'resto de lineas
                Dim L As Integer
                L = Len(Text1(0).Text)
                
                If L = 18 Then
                    Sql3 = "select * from anecoop where mid(expediente_id,1,1) > '0' and mid(expediente_id,10,9) = " & DBSet(Right(Text1(0).Text, 9), "T")
                Else
                    Sql3 = "select * from anecoop where mid(expediente_id,1,1) > '0' and mid(expediente_id, " & 18 - L + 1 & "," & L & ") =  " & DBSet(Text1(0).Text, "T")
                End If
                
                Set RS3 = New ADODB.Recordset
                RS3.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RS3.EOF
                    PNeto = DevuelveValor("select peso_neto from anecoop where expediente_id = " & DBSet(RS3!expediente_id, "T"))
                
                    vImporte1 = Round2(PNeto * vImporte / vPNeto, 2)
                    
                    ImporteTot = ImporteTot + vImporte1
                
                    'insertamos los pagos si los hay
                    Sql = "insert into anecoop_pago (expediente_id,expediente_pagoid,tipo_pago,num_factura,fecha_factura,num_liquidacion,importe,fecha_pago,fecha_pago_sc,estado) "
                    Sql = Sql & " select " & DBSet(RS3!expediente_id, "T") & ",expediente_pagoid,tipo_pago,num_factura,fecha_factura,num_liquidacion," & DBSet(vImporte1, "N") & ",fecha_pago,fecha_pago_sc,estado "
                    Sql = Sql & " from anecoop_pago where expediente_id = " & DBSet(Rs2!expediente_id, "T") & " and expediente_pagoid = " & DBSet(Rs2!expediente_pagoid, "T")
                    conn.Execute Sql
                
                    RS3.MoveNext
                Wend
                Set RS3 = Nothing
                
                ' si hay diferencias updateamos en la primera
                If ImporteTot <> vImporte Then
                    Sql = "update anecoop_pago set importe = importe + " & DBSet(vImporte - ImporteTot, "N")
                    Sql = Sql & "where expediente_id = " & DBSet(Text1(0).Text, "T")
                    Sql = Sql & " and expediente_pagoid = " & DBSet(Rs2!expediente_pagoid, "T")
                
                    conn.Execute Sql
                End If
            
                Rs2.MoveNext
            Wend
        End If
        
    End If

    CrearExpedientes = True
    conn.CommitTrans
    Exit Function

eCrearExpedientes:
    conn.RollbackTrans
    MuestraError Err.Number, "Crear Expedientes", Err.Description
End Function

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub

Private Sub mnImprimir_Click()
'Imprimir Factura
    
    If Data1.Recordset.EOF Then Exit Sub
    
    BotonImprimir
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de albaranes: scaalb
        If BLOQUEADesdeFormulario(Me) Then
            If BloqueaLineasAlb Then BotonModificarLinea (1)
        End If
         
    Else   'Modificar albaran
        'bloquea la tabla cabecera de albaranes: scaalb
        If BLOQUEADesdeFormulario(Me) Then
            BotonModificar
        End If
    End If
End Sub


Private Function BloqueaLineasAlb() As Boolean
'bloquea todas las lineas de la factura
Dim Sql As String

    On Error GoTo EBloqueaLin

    BloqueaLineasAlb = False
    'bloquear cabecera albaranes
    Sql = "select * FROM anecoop "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaLineasAlb = True

EBloqueaLin:
    If Err.Number <> 0 Then BloqueaLineasAlb = False
End Function

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub



'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String
Dim cadMen As String
Dim Sql As String
Dim Nregs As Long

        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 4, 7 'Fecha albaran y fecha liquidacion
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index), True
            
        Case 26 ' Porcentaje iva
            PonerFormatoDecimal Text1(Index), 4
            
        Case 24 ' precio
            PonerFormatoDecimal Text1(Index), 5
        
        Case 5 ' nro pedido
            PonerFormatoEntero Text1(Index)
            
        Case 22, 23 ' cajas, peso neto
            PonerFormatoEntero Text1(Index)
            
        Case 8, 9, 10, 12
            PonerFormatoDecimal Text1(Index), 1
    End Select
    
    If Index = 3 Or Index = 25 Then CargarAlbaran
    
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
Dim CadB1 As String

Dim cadAux As String
Dim I As Integer

    For I = 0 To Text1.Count - 1
        Text1(I).Tag = Replace(Text1(I).Tag, "|T|", "|TT|")
    Next I
    
    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    CadB1 = ObtenerBusqueda2(Me, BuscaChekc, 2, "FrameAux1")
 
    For I = 0 To Text1.Count - 1
        Text1(I).Tag = Replace(Text1(I).Tag, "|TT|", "|T|")
    Next I
 
 
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        
        If CadB <> "" And CadB1 <> "" Then
            MandaBusquedaPrevia CadB & " and " & CadB1
        Else
            If CadB = "" And CadB1 <> "" Then
                MandaBusquedaPrevia CadB1
            Else
                MandaBusquedaPrevia CadB
            End If
        End If
        
    ElseIf CadB <> "" Or CadB1 <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select anecoop.* from " & NombreTabla & " LEFT JOIN anecoop_pago ON anecoop.expediente_id = anecoop_pago.expediente_id "
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB
        If CadB1 <> "" And CadB <> "" Then CadenaConsulta = CadenaConsulta & " and "
        CadenaConsulta = CadenaConsulta & CadB1
        CadenaConsulta = CadenaConsulta & " GROUP BY anecoop.expediente_id " & Ordenacion
        
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
    cad = cad & "Nº.Expediente|anecoop.expediente_id|T||30·"
    cad = cad & "Linea|anecoop.linea_expediente|T||20·" 'ParaGrid(Text1(3), 10, "Cliente")
    cad = cad & "Campaña|anecoop.codigo_campanya|T||20·"
    Tabla = NombreTabla & " LEFT JOIN anecoop_pago ON anecoop.expediente_id = anecoop_pago.expediente_id "
    
    Titulo = "Expedientes"
    devuelve = "0|1|2|"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = CadB
        
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|2|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'        frmB.vConexionGrid = cAgro  'Conexión a BD: Ariagro
        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
'        If EsCabecera Then
'            PonerCadenaBusqueda
'            Text1(0).Text = Format(Text1(0).Text, "0000000")
'        End If
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(kCampo)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(kCampo)
'            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        '--monica
        'LLamaLineas Modo, 0, "DataGrid2"
        PonerCampos
    End If


    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCamposLineas()
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafac1 de la factura seleccionada
Dim b As Boolean
Dim b2 As Boolean
Dim I As Integer

    On Error GoTo EPonerLineas

    If Data1.Recordset.EOF Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    If Data1.Recordset.RecordCount > 0 Then
        CargaGrid DataGrid3, Adoaux(1), True
    Else
        CargaGrid DataGrid3, Adoaux(1), False
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single
Dim b As Boolean

    On Error Resume Next


    If Data1.Recordset.EOF Then Exit Sub
    
    b = PonerCamposForma2(Me, Data1, 2, "Frame2")
    
    'poner descripcion campos
    Modo = 4
    
    Modo = 2
    
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    CargarAlbaran
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario

    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim I As Byte, Numreg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
'    ActualizarToolbar Modo, Kmodo
    BuscaChekc = ""
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    Numreg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then Numreg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
          
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    
    b = (Modo = 4) Or (Modo = 2) Or (Modo = 0)
    'Campos Nº Albarán bloqueado y en azul
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(18), b, True
    BloquearTxt Text1(19), b, True
    
    'Si no es modo lineas Boquear los TxtAux
    For I = 1 To txtAux.Count - 1
        txtAux(I).visible = False
        BloquearTxt txtAux(I), True
    Next I
    txtAux2(2).visible = False
    
    
    
    BloquearBtn Me.btnBuscar(0), True
    BloquearBtn Me.btnBuscar(1), True
    BloquearBtn Me.btnBuscar(2), True
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
    BloquearFrameAux Me, "FrameAux1", Modo, 1
    
        
    If Modo = 1 Then
        For I = 2 To 8
    '        BloquearTxt txtAux(I), (Modo <> 1)
            txtAux(I).Enabled = True
            txtAux(I).visible = True
            txtAux(I).Locked = False
        Next I
        If Modo = 1 Then
            Dim anc As Single
              anc = DataGrid3.Top
              If DataGrid3.Row < 0 Then
                  anc = anc + 215 '210
              Else
                  anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 5
              End If
            
              LLamaLineas Modo, anc, "DataGrid3"
        End If
    End If
    
    If Modo = 2 Then LLamaLineas 0, anc, "DataGrid3"
        
    Check1(0).Enabled = (Modo = 1) Or ((Modo = 3 Or Modo = 4) And vUsu.Nivel = 0)
        
        
    ' ***************************
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean
Dim Serie As String
Dim Sql As String

    On Error GoTo EDatosOK

    DatosOk = False
    
    'comprobamos datos OK de la tabla scaalb
    b = CompForm2(Me, 2, "Frame2") ' , 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    ' Comprobamos que el albaran y numero de linea existen en albaranes
    If Modo = 3 Or Modo = 4 Then
        If Text1(3).Text <> "" And Text1(25).Text <> "" Then
            Sql = "select count(*) from albaran_variedad where numalbar = " & DBSet(Text1(3).Text, "N") & " and numlinea = " & DBSet(Text1(25).Text, "N")
            If TotalRegistros(Sql) = 0 Then
                MsgBox "No existe el albarán o la linea del albarán. Revise.", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim I As Byte

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    For I = 4 To 7
        If I = 4 Or I = 6 Or I = 7 Then
            If txtAux(I).Text = "" Then
                MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(I)
                Exit Function
            End If
        End If
    Next I
            
    DatosOkLinea = b
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

    If BloqueaRegistro(NombreTabla, "expediente_id = " & Data1.Recordset!expediente_id) Then
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        Select Case Button.Index
            Case 1
                BotonAnyadirLinea Index
            Case 2
                BotonModificarLinea Index
            Case 3
                BotonEliminarLinea Index
            Case Else
        End Select
    End If

End Sub


Private Sub BotonEliminarLinea(Index As Integer)
Dim cad As String
Dim Sql As String
Dim Mens As String
Dim b As Boolean

    On Error GoTo EEliminarLinea

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    b = True

    ' *************** canviar la pregunta ****************
    cad = "¿Seguro que desea eliminar el Pago?"
    cad = cad & vbCrLf & "Expediente: " & Adoaux(1).Recordset.Fields(0)
    cad = cad & vbCrLf & "Pago: " & Adoaux(1).Recordset.Fields(1)
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminarLinea
        Screen.MousePointer = vbHourglass
        NumRegElim = Adoaux(1).Recordset.AbsolutePosition
        
        If Not EliminarLinea Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            If SituarDataTrasEliminar(Adoaux(1), NumRegElim) Then
                PonerCampos
            Else
                PonerCampos
'                        LimpiarCampos
'                        PonerModo 0
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
       
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Or Not b Then MuestraError Err.Number, "Eliminar Linea de Factura", Err.Description

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            BotonVerTodos
        
        Case 4  'Añadir
            mnNuevo_Click

        Case 5  'Modificar
            mnModificar_Click
        Case 6  'Borrar
            mnEliminar_Click
            
        Case 8  ' desdoblar expediente
            mnDesdoblar_Click
        
        Case 9  ' Impresion de albaran
            mnImprimir_Click
        Case 11    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub ActualizarToolbar(Modo As Byte, Kmodo As Byte)
'Modo: Modo antiguo
'Kmodo: Modo que se va a poner

    If (Modo = 5) And (Kmodo <> 5) Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 3
'        Toolbar1.Buttons(5).ToolTipText = "Nuevo Albaran"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 4
        Toolbar1.Buttons(5).ToolTipText = "Modificar Factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 5
        Toolbar1.Buttons(6).ToolTipText = "Eliminar Factura"
    End If
    If Kmodo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 12
'        Toolbar1.Buttons(5).ToolTipText = "Nueva linea"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 13
        Toolbar1.Buttons(5).ToolTipText = "Modificar linea factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 14
        Toolbar1.Buttons(6).ToolTipText = "Eliminar linea factura"
    End If
End Sub
    


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Opcion As Byte
Dim Sql As String

    On Error GoTo ECargaGRid

    b = DataGrid3.Enabled
    Select Case vDataGrid.Name
        Case "DataGrid3" 'envases
            Opcion = 1
    End Select
    
    Sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    vDataGrid.RowHeight = 270
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b
    
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
    
    On Error GoTo ECargaGRid

    Select Case vDataGrid.Name
         Case "DataGrid3" 'slialb lineas de envases
'select codtipom,numfactu,fecfactu,numlinea,codartic,sartic.nomartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva
            tots = "N||||0|;N||||0|;S|txtAux(2)|T|TP|400|;S|txtAux2(2)|T|Tipo Pago|1000|;"
            tots = tots & "S|txtAux(3)|T|Factura|1500|;S|txtAux(4)|T|Fecha Fra|1500|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|txtAux(5)|T|Liquidacion|1500|;S|txtAux(6)|T|Importe|1200|;"
            tots = tots & "S|txtAux(7)|T|Fecha Pago|1500|;S|btnBuscar(1)|B|||;S|txtAux(8)|T|Fecha Pago SC|1500|;S|btnBuscar(2)|B|||;N||||0|;S|chkAux(0)|CB|GR|360|;"
            arregla tots, DataGrid3, Me
     
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
  
    ConseguirFocoLin txtAux(Index)
End Sub



Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Sql As String
Dim devuelve As String
Dim b As Boolean
Dim TipoDto As Byte
Dim vCStock As CStock

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 2 ' Tipo de Pago
            txtAux(2).Text = UCase(txtAux(2).Text)
            Select Case txtAux(2).Text
                Case "M"
                    txtAux2(2).Text = "Mercancia"
                Case "I"
                    txtAux2(2).Text = "Iva"
                Case "IM"
                    txtAux2(2).Text = "Iva Manual"
                Case Else
                    txtAux2(2).Text = ""
            End Select
            
        Case 4, 7, 8 'fecha factura
            If PonerFormatoFecha(txtAux(Index)) Then
                If Index = 8 Then Me.cmdAceptar.SetFocus
            End If
        
        Case 5 'nro liquidacion
            PonerFormatoEntero txtAux(Index)
            
        Case 6 ' importe
            PonerFormatoDecimal txtAux(Index), 1
            
    End Select
    
End Sub




Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String, Sql2 As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    'Eliminar en tablas de cabecera de albaran
    '------------------------------------------
    Sql = " " & ObtenerWhereCP(True)
    
    'Lineas de envases (slialb)
    conn.Execute "Delete from anecoop_pago where expediente_id = " & DBSet(Text1(0).Text, "T")
    
    'Cabecera de factura
    conn.Execute "Delete from " & NombreTabla & Sql
        
    b = True

FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Expediente", Err.Description & " " & Mens
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Function EliminarLinea() As Boolean
Dim Sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String
Dim vCStock As CStock

    On Error GoTo FinEliminar

    b = False
    If Adoaux(1).Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""
    
    'Eliminar en tablas de pagos
    '------------------------------------------
    Sql = " where expediente_id = " & DBSet(Adoaux(1).Recordset.Fields(0), "T")
    Sql = Sql & " and expediente_pagoid = " & DBSet(Adoaux(1).Recordset.Fields(1), "T")

    'Lineas pagos
    conn.Execute "Delete from anecoop_pago " & Sql
    b = True
    
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Pagos del Expediente", Err.Description & " " & Mens
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        EliminarLinea = False
    Else
        conn.CommitTrans
        EliminarLinea = True
    End If
End Function

Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    CargaGrid DataGrid3, Me.Adoaux(1), False 'pagos
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
             PonerModo 2
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = "expediente_id= " & DBSet(Text1(0).Text, "T") & " and linea_expediente = " & DBSet(Text1(18).Text, "T") & " and codigo_campanya = " & DBSet(Text1(19).Text, "T")
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
    
    Select Case Opcion
        Case 1  'pagos
            Sql = "SELECT expediente_id, expediente_pagoid, tipo_pago, CASE tipo_pago WHEN ""M"" THEN ""Mercancia"" WHEN ""I"" THEN ""Iva"" WHEN ""IM"" THEN ""Iva Manual"" END dtipo_pago, num_factura, fecha_factura, num_liquidacion, importe, fecha_pago, fecha_pago_sc, "
            Sql = Sql & "idcontab, IF(idcontab=1,'*','') as didcontab "
            Sql = Sql & " FROM anecoop_pago "
            Sql = Sql & " WHERE (1=1) "
    End Select
    
    If enlaza Then
        Sql = Sql & " and expediente_id= " & DBSet(Text1(0).Text, "T")
    Else
        Sql = Sql & " and expediente_id = -1"
    End If
    Sql = Sql & " ORDER BY expediente_pagoid"
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim I As Integer

        b = ((Modo = 2) Or (Modo = 0)) 'Or (Modo = 5 And ModificaLineas = 0)
        'Buscar
        Toolbar1.Buttons(1).Enabled = b
        Me.mnBuscar.Enabled = b
        'Vore Tots
        Toolbar1.Buttons(2).Enabled = b
        Me.mnVerTodos.Enabled = b
        'Añadir
        Toolbar1.Buttons(4).Enabled = b
        Me.mnModificar.Enabled = b
        
        
        b = (Modo = 2 And Data1.Recordset.RecordCount > 0)
        'Modificar
        Toolbar1.Buttons(5).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(6).Enabled = b
        Me.mnEliminar.Enabled = b
        'Impresión de factura
        Toolbar1.Buttons(8).Enabled = (Modo = 2)
        Me.mnImprimir.Enabled = (Modo = 2)
        

    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 2)
    For I = 1 To 1
        ToolAux(I).Buttons(1).Enabled = b
        
        If b Then
            Select Case I
              Case 0
                bAux = (b And Me.Adoaux(0).Recordset.RecordCount > 0)
              Case 1
                bAux = (b And Me.Adoaux(1).Recordset.RecordCount > 0)
            End Select
        End If
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I


End Sub


Private Sub BotonImprimir()
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar un Albarán para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 29 'Impresion de albaran de Envases
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de factura
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Nº Albaran
        devuelve = "{" & NombreTabla & ".numalbar}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "numalbar = " & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    End If
    
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Impresión de Albarán Envases"
            .ConSubInforme = True
            .Show vbModal
    End With
End Sub

'Private Sub TxtAux3_GotFocus(Index As Integer)
'    ConseguirFoco txtAux3(Index), Modo
'End Sub
'
'Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
'End Sub
'
'Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
'
'Private Sub TxtAux3_LostFocus(Index As Integer)
'Dim TipoDto As Byte
'Dim ImpDto As String
'Dim Unidades As String
'Dim cantidad As String
'Dim cad As String
'
'    'Quitar espacios en blanco
'    If Not PerderFocoGnralLineas(txtAux3(Index), ModificaLineas) Then Exit Sub
'
'    Select Case Index
'        Case 4 'Albaran
'            If txtAux3(Index) <> "" Then PonerFormatoEntero txtAux3(Index)
'
'            CargarDatosAlbaran txtAux3(4).Text, txtAux3(5).Text
'
'        Case 5 'Linea de albaran
'            If txtAux3(Index) <> "" Then PonerFormatoEntero txtAux3(Index)
'
'            If txtAux3(4).Text <> "" And txtAux3(5).Text <> "" Then
'                If AlbaranFacturado(txtAux3(4).Text, txtAux3(5).Text) Then
'                    cad = "Esta línea de Albarán está facturada. " & vbCrLf & vbCrLf & "    ¿ Desea continuar ? "
'                    If MsgBox(cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                        CargarDatosAlbaran txtAux3(4).Text, txtAux3(5).Text
'                    Else
'                        txtAux3(4).Text = ""
'                        txtAux3(5).Text = ""
'                    End If
'                Else
'                    CargarDatosAlbaran txtAux3(4).Text, txtAux3(5).Text
'                End If
'            End If
'
'            If txtAux3(4).Text = "" Or txtAux3(5).Text = "" Then
'                PonerFoco txtAux3(4)
'            Else
'                PonerFoco txtAux3(8)
'            End If
'
'        Case 8 'precio bruto
'            If txtAux3(Index).Text <> "" Then
'                If PonerFormatoDecimal(txtAux3(Index), 7) Then
'
'                    Select Case TipoFacturarForfaits(txtAux3(4).Text, txtAux3(5).Text)
'                        Case 0  'por unidades
'                            txtAux3(10).Text = Round2(CCur(ImporteSinFormato(txtAux3(Index).Text)) * CCur(ImporteSinFormato(txtAux3(15).Text)), 2)
'                            PonerFormatoDecimal txtAux3(10), 3
'                        Case 1  'por kilos
'                            txtAux3(10).Text = Round2(CCur(ImporteSinFormato(txtAux3(Index).Text)) * CCur(ImporteSinFormato(txtAux3(6).Text)), 2)
'                            PonerFormatoDecimal txtAux3(10), 3
'                        Case Else
'
'                    End Select
'
'                    cmdAceptar.SetFocus
'                Else
'                    Exit Sub
'                End If
'            End If
'
'        Case 10 'importe bruto
'            If txtAux3(Index).Text <> "" Then
'                If PonerFormatoDecimal(txtAux3(Index), 3) Then
'
'                    Select Case TipoFacturarForfaits(txtAux3(4).Text, txtAux3(5).Text)
'                        Case 0
'                            Unidades = ComprobarCero(txtAux3(15).Text)
'                            If CCur(Unidades) <> 0 Then
'                                txtAux3(8).Text = Round2(CCur(ImporteSinFormato(txtAux3(Index).Text)) / CCur(Unidades), 4)
'                            Else
'                                txtAux3(8).Text = 0
'                            End If
'                            PonerFormatoDecimal txtAux3(8), 7
'                        Case 1
'                            cantidad = ComprobarCero(txtAux3(6).Text)
'                            If CCur(cantidad) <> 0 Then
'                                txtAux3(8).Text = Round2(CCur(ImporteSinFormato(txtAux3(Index).Text)) / CCur(cantidad), 4)
'                            Else
'                                txtAux3(8).Text = 0
'                            End If
'                            PonerFormatoDecimal txtAux3(8), 7
'                        Case Else
'
'                    End Select
'
'                    cmdAceptar.SetFocus
'               Else
'                    Exit Sub
'               End If
'            End If
'    End Select
'
'If ((Index = 8 And txtAux3(Index).Text <> "") Or (Index = 10 And txtAux3(Index).Text <> "")) Then
'        Dim Campo2 As String
'        Campo2 = "nrodecprec"
'        TipoDto = DevuelveDesdeBDNew(cAgro, "clientes", "tipodtos", "codclien", Text1(3).Text, "N", Campo2)
'        Select Case TipoFacturarForfaits(txtAux3(4).Text, txtAux3(5).Text)
'            Case 0 ' unidades
''                ImpDto = CalcularImporteDto(txtAux3(15).Text, txtAux3(8).Text, txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!impdtoc, "N")), False)
''                txtAux3(11).Text = CalcularImporte(txtAux3(15).Text, txtAux3(8).Text, txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto)
'                Unidades = ComprobarCero(txtAux3(15).Text)
'                ImpDto = CalcularImporteDto(txtAux3(15).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(Unidades)), txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!impdtoc, "N")), False)
'                txtAux3(11).Text = CalcularImporte(txtAux3(15).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(Unidades)), txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto, txtAux3(10).Text)
'                PonerFormatoDecimal txtAux3(11), 1
'
'                'precio neto
'                If ComprobarCero(txtAux3(15).Text) <> "0" Then
'                    txtAux3(9).Text = Round2(CCur(ImporteSinFormato(txtAux3(11).Text)) / CCur(ImporteSinFormato(txtAux3(15).Text)), CCur(Campo2))
'                End If
'                PonerFormatoDecimal txtAux3(9), 7
'
'            Case 1 ' kilos
''                ImpDto = CalcularImporteDto(txtAux3(6).Text, txtAux3(8).Text, txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!impdtoc, "N")), False)
''                txtAux3(11).Text = CalcularImporte(txtAux3(6).Text, txtAux3(8).Text, txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto)
'                cantidad = ComprobarCero(txtAux3(6).Text)
'                ImpDto = CalcularImporteDto(txtAux3(6).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(cantidad)), txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!impdtoc, "N")), False)
'                txtAux3(11).Text = CalcularImporte(txtAux3(6).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(cantidad)), txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto, txtAux3(10).Text)
'                PonerFormatoDecimal txtAux3(11), 1
'
'                'precio neto
'                If ComprobarCero(txtAux3(6).Text) <> "0" Then
'                    txtAux3(9).Text = Round2(CCur(ImporteSinFormato(txtAux3(11).Text)) / CCur(ImporteSinFormato(txtAux3(6).Text)), CCur(Campo2))
'                End If
'                PonerFormatoDecimal txtAux3(9), 7
'
'            Case Else
'
'        End Select
'
'    End If
'
'End Sub

Private Sub CargaCombo()
Dim RS As ADODB.Recordset
Dim Sql As String
Dim I As Byte
    
'    Combo1(0).Clear
'
'    Combo1(0).AddItem "Normal"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
'
'    Combo1(0).AddItem "Exento"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
'
'    Combo1(0).AddItem "Recargo Equiv."
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
End Sub

Private Function ModificaCabecera() As Boolean
Dim b As Boolean
Dim MenError As String

    On Error GoTo EModificarCab

    conn.BeginTrans
    
'    MenError = "Eliminando Costes"
'    b = EliminarCostes(Data1.Recordset.Fields(0))
    
    b = ModificaDesdeFormulario2(Me, 2, "Frame2")
    
    
    MenError = "Recalculando Importes Netos de lineas"

EModificarCab:
    If Err.Number <> 0 Then
        MenError = "Modificando Albaran Envases." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        b = False
    End If
    If b Then
        ModificaCabecera = True
        conn.CommitTrans
    Else
        ModificaCabecera = False
        conn.RollbackTrans
    End If
End Function


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    On Error GoTo EInsertarCab
    
    Sql = CadenaInsertarDesdeForm(Me)
    If Sql <> "" Then
        If InsertarOferta(Sql, vTipoMov) Then
            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
            PonerCadenaBusqueda
            PonerModo 2
            'Ponerse en Modo Insertar Lineas
'                BotonMtoLineas 0, "Variedades"
            BotonAnyadirLinea 0
        End If
    End If
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Albaranes
    'para ello vemos si existe una factura con ese contador y si existe la incrementamos
    vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla de Expedientes Anecoop (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    
EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarOferta = True
        Else
            conn.RollbackTrans
            InsertarOferta = False
        End If
End Function


'Private Sub CargaForaGrid()
'    If DataGrid2.Columns.Count <= 2 Then Exit Sub
'    ' *** posar als camps de fora del grid el valor de la columna corresponent ***
'    Text3(0) = DataGrid2.Columns(12).Text    'Fecha
'    Text3(1) = DataGrid2.Columns(13).Text    'Matricula
'    Text3(2) = DataGrid2.Columns(14).Text    'Destino
'    Text3(3) = DataGrid2.Columns(15).Text   'Variedad
'    Text3(4) = DataGrid2.Columns(16).Text   'Confeccion
'    ' *** Si fora del grid n'hi han camps de descripció, posar-los valor ***
'    ' **********************************************************************
'End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean
Dim Mens As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
'        Case 0: nomFrame = "FrameAux0" 'variedades
    nomframe = "FrameAux1" 'envases
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        If InsertarLineaEnv(txtAux(1).Text) Then
            b = BloqueaRegistro("anecoop", "expediente_id = " & DBSet(Data1.Recordset!expediente_id, "T"))
            CargaGrid DataGrid3, Adoaux(1), True
            If b Then BotonAnyadirLinea 1
        End If
    End If
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim I As Integer
    
    ModificaLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModificaLineas = 0 Then Exit Sub
    End If
       
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    BloquearTxt Text1(18), True
    BloquearTxt Text1(19), True
    
    
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    vTabla = "anecoop_pago"
    ' ********************************************************
    
    vWhere = "expediente_id = " & DBSet(Text1(0).Text, "T")
    
    NumF = SugerirCodigoSiguienteStr(vTabla, "expediente_pagoid", vWhere)
    ' ***************************************************************

    AnyadirLinea DataGrid3, Adoaux(1)

    anc = DataGrid3.Top
    If DataGrid3.Row < 0 Then
        anc = anc + 215 '210
    Else
        anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 5
    End If
  
    LLamaLineas ModificaLineas, anc, "DataGrid3"

    LimpiarCamposLin "FrameAux1"
    txtAux(0).Text = Text1(0).Text 'numexpediente
    txtAux(1).Text = NumF
    PonerFoco txtAux(2)
    
    BloquearBtn Me.btnBuscar(0), False
    BloquearBtn Me.btnBuscar(1), False
    BloquearBtn Me.btnBuscar(2), False
' ******************************************
End Sub

Private Sub LimpiarCamposLin(frameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, frameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As String
Dim cad As String
Dim Sql As String
Dim vCStock As CStock
Dim b As Boolean
Dim Mens As String
    
    On Error GoTo eModificarLinea

    ModificarLinea = False
    Sql = ""

    ' *** posa els noms del frames, tant si son de grid com si no ***
    nomframe = "FrameAux1" 'envases
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        
        If DatosOkLineaEnv() Then
            '#### LAURA 15/11/2006
            conn.BeginTrans
            
            Sql = "UPDATE anecoop_pago Set tipo_pago = " & DBSet(txtAux(2).Text, "T") & ", num_factura=" & DBSet(txtAux(3).Text, "T") & ", "
            Sql = Sql & "fecha_factura=" & DBSet(txtAux(4).Text, "F") & ", "
            Sql = Sql & "num_liquidacion= " & DBSet(txtAux(5).Text, "N") & ", "
            Sql = Sql & "importe= " & DBSet(txtAux(6).Text, "N") & ", " 'importe
            Sql = Sql & "fecha_pago= " & DBSet(txtAux(7).Text, "F")
            Sql = Sql & " where expediente_id = " & DBSet(Adoaux(1).Recordset!expediente_id, "T") & " AND expediente_pagoid=" & Adoaux(1).Recordset!expediente_pagoid
            conn.Execute Sql
        
        End If
            
        ModificaLineas = 0
        
        
        V = Adoaux(1).Recordset.Fields(1) 'el 2 es el nº de llinia
        CargaGrid DataGrid3, Adoaux(1), True

        ' *** si n'hi han tabs ***
'            SSTab1.Tab = 1

        DataGrid3.SetFocus
        Adoaux(1).Recordset.Find (Adoaux(1).Recordset.Fields(1).Name & " =" & DBSet(V, "T"))

        LLamaLineas ModificaLineas, 0, "DataGrid3"

        b = True

    End If
        
eModificarLinea:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description & vbCrLf & Mens
        b = False
    End If
    
    If b Then
        conn.CommitTrans
        ModificarLinea = True
    Else
        conn.RollbackTrans
        ModificarLinea = False
    End If
End Function
        

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
Dim Cliente As String

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    
    DatosOkLlin = b

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Function InsertarLineaEnv(numlinea As String) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
Dim Sql As String
Dim vWhere As String
Dim b As Boolean
Dim DentroTRANS As Boolean

    InsertarLineaEnv = False
    Sql = ""
    DentroTRANS = False
    
    'Conseguir el siguiente numero de linea
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
'    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
    
    If DatosOkLineaEnv() Then 'Lineas de factura
        'Inserta en tabla "facturas_envases"
        Sql = "INSERT INTO anecoop_pago "
        Sql = Sql & "(expediente_id, expediente_pagoid, tipo_pago, num_factura, fecha_factura, num_liquidacion, importe, fecha_pago, fecha_pago_sc ) "
        Sql = Sql & "VALUES (" & DBSet(txtAux(0).Text, "T") & ", " & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T", "S") & ","
        Sql = Sql & DBSet(txtAux(3).Text, "T", "S") & ", "
        Sql = Sql & DBSet(txtAux(4).Text, "F", "S") & ", "
        Sql = Sql & DBSet(txtAux(5).Text, "N", "S") & ", " & DBSet(txtAux(6).Text, "N", "S") & ", "
        Sql = Sql & DBSet(txtAux(7).Text, "F", "S") & ","
        Sql = Sql & DBSet(txtAux(8).Text, "F", "S") & ")"
     Else
        Exit Function
     End If
    
    If Sql <> "" Then
        On Error GoTo EInsertarLineaEnv
        conn.BeginTrans
        DentroTRANS = True
        
        'insertar la linea
        conn.Execute Sql
        
        'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
        'en actualizar stock comprobamos si el articulo tiene control de stock
        
        b = True
        
    End If
    
    If b Then
        conn.CommitTrans
        InsertarLineaEnv = True
    Else
        conn.RollbackTrans
         InsertarLineaEnv = False
    End If
    Exit Function
    
EInsertarLineaEnv:
    If Err.Number <> 0 Then
        InsertarLineaEnv = False
        If DentroTRANS Then conn.RollbackTrans
        MuestraError Err.Number, "Insertar Lineas Expedientes" & vbCrLf & Err.Description
'        b = False
    End If
'    If b Then
'        Conn.CommitTrans
'        InsertarLinea = True
'    Else
'        Conn.RollbackTrans
'         InsertarLinea = False
'    End If
End Function


Private Function DatosOkLineaEnv() As Boolean
Dim b As Boolean
Dim I As Byte
    
    On Error GoTo EDatosOkLineaEnv

    DatosOkLineaEnv = False
    b = True

    DatosOkLineaEnv = b
    
EDatosOkLineaEnv:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub CargarAlbaran()
Dim Sql As String
Dim RS As ADODB.Recordset

    On Error Resume Next

    Text3(17).Text = ""
    Text3(16).Text = ""
    Text3(14).Text = ""
    Text3(13).Text = ""
    Text3(23).Text = ""
    Text3(22).Text = ""

    If Text1(3).Text = "" Or Text1(25).Text = "" Then Exit Sub

    Sql = "SELECT albaran_variedad.numalbar, numlinea, albaran_variedad.codvarie, a.nomvarie as nomvarie1, albaran_variedad.codvarco, "
    Sql = Sql & " b.nomvarie as nomvarie2, albaran_variedad.codmarca, marcas.nommarca, albaran_variedad.codforfait, forfaits.nomconfe, "
    Sql = Sql & " categori, pesobrut, totpalet, preciopro, numcajas, unidades, pesoneto " ', preciodef, albaran_variedad.codincid, inciden.nomincid, "
    Sql = Sql & ", albaran_variedad.codpalet, preciodef "
    Sql = Sql & " FROM albaran_variedad, variedades a, variedades b, marcas, forfaits, inciden " 'lineas de variedades del albaran
    Sql = Sql & " WHERE albaran_variedad.numalbar = " & DBSet(Text1(3).Text, "N") & " and albaran_variedad.numlinea = " & DBSet(Text1(25).Text, "N")
    Sql = Sql & " and albaran_variedad.codvarie = a.codvarie "
    Sql = Sql & " and albaran_variedad.codvarco = b.codvarie"
    Sql = Sql & " and albaran_variedad.codmarca = marcas.codmarca "
    Sql = Sql & " and albaran_variedad.codforfait = forfaits.codforfait "
    Sql = Sql & " and albaran_variedad.codincid = inciden.codincid "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        Text3(17).Text = DBLet(RS!nomvarie1)
        Text3(16).Text = DBLet(RS!nommarca)
        Text3(14).Text = DBLet(RS!nomconfe)
        Text3(13).Text = DBLet(RS!categori)
        Text3(23).Text = Format(DBLet(RS!Pesoneto), "###,###,##0")
        Text3(22).Text = Format(DBLet(RS!NumCajas), "###,###,##0")
    End If

    Set RS = Nothing

End Sub
