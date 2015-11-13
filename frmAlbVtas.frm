VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAlbVtas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Albaranes Venta"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "frmAlbVtas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   9
      Left            =   4200
      MaxLength       =   12
      TabIndex        =   4
      Tag             =   "Referencia|T|S|||albaran_variedad|referencia|||"
      Text            =   "Ref.Lin"
      Top             =   4680
      Width           =   630
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   8
      Left            =   6750
      MaxLength       =   30
      TabIndex        =   7
      Tag             =   "Imp.Neto|N|N|||facturas_variedad|impornet|##,###,##0.00|N|"
      Text            =   "ImpFact"
      Top             =   4680
      Width           =   540
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   7
      Left            =   6240
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
      Text            =   "Fac"
      Top             =   4680
      Width           =   450
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   6
      Left            =   5550
      MaxLength       =   30
      TabIndex        =   17
      Text            =   "*"
      Top             =   4680
      Width           =   585
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   4890
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "Peso Neto|N|S|||albaran_variedad|pesoneto|###,##0|N|"
      Text            =   "pneto"
      Top             =   4710
      Width           =   540
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   3060
      MaxLength       =   30
      TabIndex        =   16
      Text            =   "Variedad"
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   2475
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Num.Linea|N|N|||albaran_variedad|numlinea|00|S|"
      Text            =   "lin"
      Top             =   4680
      Width           =   540
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   3600
      MaxLength       =   12
      TabIndex        =   3
      Tag             =   "Referencia Cl|T|S|||albaran|refclien|||"
      Text            =   "Ref"
      Top             =   4680
      Width           =   540
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   120
      MaxLength       =   10
      TabIndex        =   0
      Tag             =   "Numero Albarán|N|N|||albaran|numalbar|000000|S|"
      Text            =   "numalb"
      Top             =   4680
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   960
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Fecha albaran|F|N|||albaran|fechaalb|||"
      Text            =   "fecalb"
      Top             =   4680
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6705
      TabIndex        =   8
      Top             =   5340
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7965
      TabIndex        =   9
      Top             =   5340
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7965
      TabIndex        =   12
      Top             =   5310
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   5175
      Width           =   2385
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
         Height          =   255
         Left            =   40
         TabIndex        =   11
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   3120
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlbVtas.frx":000C
      Height          =   4410
      Left            =   120
      TabIndex        =   13
      Top             =   540
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   7779
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
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
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmAlbVtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funció BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funció BotonBuscar() canviar el nom de la clau primaria
' 5. En la funció BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funció posamaxlength() repasar el maxlength de TextAux(0)
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar algún) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada botó per a que corresponguen
' 9. En la funció CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar ademés els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funció DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funció SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' ********************************************************************************

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'altre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
'Public NuevoCodigo As String

'codigo que tiene el campo en el momento que se llama desde otro formulario
'nos situamos en ese valor
Public CodigoActual As String

Private CadenaConsulta As String
Private CadB As String

Dim Modo As Byte
'----------- MODOS ----------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'-----------------------------------------------
Dim PrimeraVez As Boolean
Dim i As Integer


Private Sub PonerModo(vModo)
Dim b As Boolean

    Modo = vModo
'    PonerIndicador lblIndicador, Modo
    b = (Modo = 2)
    
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = Not b
    Next i
    
    txtAux(4).visible = False
    txtAux(6).visible = False
    
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor.
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu 'Activar/Desact botones de menu segun permisos del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim b As Boolean

    b = (Modo = 2)
'    mnOpciones.Enabled = b
    'Buscar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(3).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = False
    Me.mnNuevo.Enabled = False
    
'    b = (b And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(7).Enabled = False
    Me.mnModificar.Enabled = False
    'Eliminar
    Toolbar1.Buttons(8).Enabled = False
    Me.mnEliminar.Enabled = False
    
    'Imprimir
    Toolbar1.Buttons(11).Enabled = b
End Sub


Private Sub BotonVerTodos()
    CargaGrid "codclien=" & CodigoActual
    CadB = ""
    PonerModo 2
End Sub


Private Sub BotonBuscar()
'    lblIndicador.Caption = "BUSQUEDA"
    ' ***************** canviar per la clau primaria ********
    CargaGrid "codclien= -1"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i

    LLamaLineas DataGrid1.Top + 206, 1
    PonerFoco txtAux(0)
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    For i = 0 To txtAux.Count - 1
        If i <> 4 And i <> 6 Then txtAux(i).Top = alto
    Next i
End Sub

Private Sub cmdAceptar_Click()
    Select Case Modo
        Case 1  'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                PonerModo 2
                CargaGrid CadB & " AND codclien = " & CodigoActual
'                lblIndicador.Caption = "RESULTADO BUSQUEDA"
                PonerFocoGrid Me.DataGrid1
            End If
    End Select
End Sub


Private Sub cmdCancelar_Click()
'On Error Resume Next

    Select Case Modo
        Case 1 'BUSQUEDA
            If CadB <> "" Then
                 CargaGrid CadB & " AND codclien = " & CodigoActual
'                lblIndicador.Caption = "RESULTADO BUSQUEDA"
            Else
                 CargaGrid "codclien = " & CodigoActual
'                lblIndicador.Caption = ""
            End If
    End Select
    PonerModo 2
    PonerFocoGrid Me.DataGrid1
'    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String
Dim i As Integer
Dim j As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    i = 0
    Do
        j = i + 1
        i = InStr(j, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, j, i - j)
            j = Val(Aux)
            cad = cad & adodc1.Recordset.Fields(j) & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     PonerContRegIndicador
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    'Ponemos el foco
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        PonerFoco txtAux(1)
'    End If

     If PrimeraVez Then
        PrimeraVez = False
'        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'            BotonAnyadir
'        Else
            PonerModo 2
            If Me.CodigoActual <> "" Then
                SituarData Me.adodc1, "codclien = " & DBSet(CodigoActual, "N"), "", True
            End If
'        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        .Buttons(2).Image = 1   'Buscar
        .Buttons(3).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(6).Image = 3   'Insertar
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        'el 9 i el 10 son separadors
'        .Buttons(11).Image = 10  'Imprimir
        .Buttons(11).Image = 11  'Salir
    End With

    '## A mano
    chkVistaPrevia.Value = CheckValueLeer(Name)
      
'    PonerOpcionesMenu  'En funcion del usuario
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "Select albaran.numalbar, albaran_variedad.numlinea, fechaalb, albaran.refclien, albaran_variedad.referencia, variedades.nomvarie, albaran_variedad.pesoneto, if(facturas_variedad.numalbar <> 0,'*',''), facturas_variedad.numfactu, facturas_variedad.impornet "
    CadenaConsulta = CadenaConsulta & " from (((albaran inner join albaran_variedad on albaran.numalbar = albaran_variedad.numalbar) inner join variedades on albaran_variedad.codvarie = variedades.codvarie) left join facturas_variedad on albaran_variedad.numalbar = facturas_variedad.numalbar and albaran_variedad.numlinea = facturas_variedad.numlinealbar)"
    '************************************************************************
    
    CadB = ""
    CargaGrid "codclien = " & CodigoActual

End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    If Modo = 4 Then TerminaBloquear
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2
                BotonBuscar
        Case 3
                BotonVerTodos
'        Case 6
'                BotonAnyadir
'        Case 7
'                mnModificar_Click
'        Case 8
'                BotonEliminar
'        Case 11 'Imprimir
'            AbrirListado (2)  'OpcionListado=2
        Case 11 'Salir
                mnSalir_Click
    End Select
End Sub


Private Sub CargaGrid(Optional vSQL As String)
    Dim Sql As String, tots As String
    
    adodc1.ConnectionString = conn
    If vSQL <> "" Then
        Sql = CadenaConsulta & " WHERE " & vSQL
    Else
        Sql = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    Sql = Sql & " ORDER BY numalbar"
    '**************************************************************++
    
    adodc1.RecordSource = Sql
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    Set DataGrid1.DataSource = adodc1
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    'numalbar, fechaalb, matriveh
    tots = "S|txtAux(0)|T|Albarán|800|;S|txtAux(3)|T|Lin.|400|;S|txtAux(1)|T|Fec.Alb.|1000|;S|txtAux(2)|T|Referencia|1200|;S|txtAux(9)|T|Ref.Lin|1150|;S|txtAux(4)|T|Variedad|1350|;S|txtAux(5)|T|P.Neto|700|;S|txtAux(6)|T|S|400|;"
    tots = tots & "S|txtAux(7)|T|Factura|800|;S|txtAux(8)|T|Importe|900|;"
    arregla tots, DataGrid1, Me
   
    DataGrid1.Columns(6).Alignment = dbgRight
    DataGrid1.Columns(7).Alignment = dbgCenter
    DataGrid1.Columns(8).Alignment = dbgCenter
    DataGrid1.Columns(9).Alignment = dbgRight

'   'Habilitamos modificar y eliminar
'   Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
'   Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
'   mnModificar.Enabled = Not adodc1.Recordset.EOF
'   mnEliminar.Enabled = Not adodc1.Recordset.EOF
'
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0, 3, 7, 5 'albaran, linea, factura, pesoneto
             PonerFormatoEntero txtAux(Index)
             
        Case 1 ' ponerformato entero
            If txtAux(Index).Text <> "" Then PonerFormatoFecha txtAux(Index)
        
        Case 8 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 3 'Tipo 3: Decimal(10,2)
            
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


Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
'End Sub
