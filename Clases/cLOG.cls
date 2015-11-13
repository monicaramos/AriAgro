VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'ACCIONES DISPONIBLES
'
'       1 .- Eliminar facturas
'       2 .- Facturacion
'       3 .- Contabilizar facturas
'       4 .- Correccion precios
'       5 .- Facturacion desde hasta COMPRAS
'       6 .- Eliminar Albaran COMPRAS
'       7 .- CCOSTES Cambio de hora de inicio en las fichadas de trabajadores a 07:00
'       8 .- CCOSTES Cambio en las fichadas, introducimos el tiempo de almuerzo que no estaba
'       9 .- Modificacion Factura Cliente
'      10 .- Modificacion Albaran de Cliente
'      11 .- Insercion de Expedientes de Anecoop

Public Function Insertar(Accion As Byte, ByRef ElUsuario As Usuario, Descripcion As String) As Boolean
Dim Sql As String

    On Error GoTo EI
    Insertar = False
    
    Sql = "insert into `slog` (`fecha`,`accion`,`usuario`,`pc`,`descripcion`,`procedencia`) values ( "
    Sql = Sql & " now()," & Accion & " ,'" & DevNombreSQL(ElUsuario.Login) & "','"
    Sql = Sql & DevNombreSQL(ElUsuario.PC) & "',"
    If Descripcion = "" Then
        Sql = Sql & "NULL,0)"
    Else
        Sql = Sql & "'" & DevNombreSQL(Descripcion) & "',0)"
    End If
    Conn.Execute Sql
    
    
    Insertar = True
    Exit Function
EI:
    MuestraError Err.Number, Err.Description

End Function




Public Function DevuelveAcciones(ByRef ListaAcciones As Collection) As Boolean
 On Error GoTo EDevuelveAcciones
    ListaAcciones.Add "1|Eliminar facturas|"
    ListaAcciones.Add "2|Facturación|"
    ListaAcciones.Add "3|Contabilizar facturas|"
    ListaAcciones.Add "4|Corrección precios|"
    ListaAcciones.Add "5|Facturacion COMPRAS|"
    ListaAcciones.Add "6|Eliminar Albarán COMPRAS|"
    ListaAcciones.Add "7|CC Cambio Hora Inicio|"
    ListaAcciones.Add "8|CC Cambios por Almuerzo|"
    ListaAcciones.Add "9|Mod.Factura Cliente|"
    ListaAcciones.Add "10|Mod.Albarán Cliente|"
    ListaAcciones.Add "11|Insertar Exp.Anecoop|"
    
    DevuelveAcciones = True
    Exit Function
EDevuelveAcciones:
    MuestraError Err.Number
End Function

