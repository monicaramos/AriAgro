Attribute VB_Name = "bus"
'NOTA: en este m�dul, adem�s, n'hi han funcions generals que no siguen de formularis (molt b�)
Option Explicit

'Definicion Conexi�n a BASE DE DATOS
'---------------------------------------------------
'Conexi�n a la BD Ariagro de la empresa
Public Conn As ADODB.Connection

'Conexi�n a la BD de Usuarios
'Public ConnUsuarios As ADODB.Connection

'Conexi�n a la BD de Contabilidad de la empresa conectada
Public ConnConta As ADODB.Connection

'Conexi�n a la BD de Contabilidad de otra empresa distinta a la conectada
Public ConnAuxCon As ADODB.Connection

'Conexi�n a la BD de Aridoc de la empresa conectada
Public ConnAridoc As ADODB.Connection

Public ConnMB As ADODB.Connection

'Que conexion a base de datos se va a utilizar
Public Const cAgro As Byte = 1 'trabajaremos con conn (conexion a BD Ariagro)
Public Const cConta As Byte = 2 'trabajaremos con connConta (cxion a BD Contabilidad)
Public Const cAridoc As Byte = 3 'trabajaremos con connAridoc (cxion a BD Aridoc)

'LOG de acciones relevantes
Public LOG As cLOG   'Se instancia , se ejecuta LOG.insertar y se elimina :LOG=nothing   Ver ejemplo borre facturas


'Definicion de clases de la aplicaci�n
'-----------------------------------------------------
Public vEmpresa As Cempresa  'Los datos de la empresa
Public vParam As Cparametros  'Parametros Generales de la Empresa (nombre, direc.,...
Public vParamAplic As CParamAplic   'parametros de la aplicacion
Public vSesion As CSesion   'Los datos del usuario que hizo login

Public vUsu As Usuario  'Datos usuario
Public vConfig As Configuracion

Public miRsAux As ADODB.Recordset

Public Const vbFPTransferencia = 1

'Definicion de FORMATOS
'---------------------------------------------------
Public FormatoFecha As String
Public FormatoFechaHora As String
Public FormatoHora As String
Public FormatoImporte As String 'Decimal(12,2)
Public FormatoPrecio As String 'Decimal(8,3)
Public FormatoCantidad As String 'Decimal(10,2)
Public FormatoPorcen As String 'Decimal(5,2) 'Porcentajes
Public FormatoExp As String  'Expedientes
Public FormatoDescuento As String 'Decimal(4,2)

Public FormatoDec10d2 As String 'Decimal(10,2)
Public FormatoDec10d3 As String 'Decimal(10,3)
Public FormatoDec5d4 As String 'Decimal(5,4)
Public FormatoDec8d4 As String 'Decimal(8,4)
Public FormatoDec6d4 As String 'Decimal(6,4)
Public FormatoDec8d2 As String 'Decimal(8,2)

Public FIni As String
Public FFin As String

'Public FormatoKms As String 'Decimal(8,4)


Public teclaBuscar As Integer 'llamada desde prismaticos

Public CadenaDesdeOtroForm As String

'Global para n� de registro eliminado
Public NumRegElim  As Long

'publica para almacenar control cambios en registros de formularios
'se utiliza en InsertarCambios
Public CadenaCambio As String
Public ValorAnterior As String

Public Confeccion As String

Public MensError As String
Public FormularioOK As String

'Para algunos campos de texto suletos controlarlos
'Public miTag As CTag

'Variable para saber si se ha actualizado algun asiento
'Public AlgunAsientoActualizado As Boolean
'Public TieneIntegracionesPendientes As Boolean

'Public miRsAux As ADODB.Recordset

Public AnchoLogin As String  'Para fijar los anchos de columna

' **** DATOS DEL LOGIN ****
'Public CodEmple As Integer
'Public codAgenc As Integer
'Public codEmpre As Integer
'Public codGrupo As Integer
'Public claEmpre As Integer
'Public TipEmple As Integer
'Public areEmple As Integer
' *************************

Public dbAriagro As BaseDatos ' base de datos para la grabacion del chivato

Public ardDB As BaseDatos ' este es la base de datos que soportar� aridoc

'[Monica]17/06/2013: para el informe de albaranes/facturas comercial
Public CategoriaValorNulo As Boolean
Public SeleccionadosTodos As Boolean

'Inicio Aplicaci�n
Public Sub Main()

'     If App.PrevInstance Then
'        MsgBox "Ariagro ya se esta ejecutando", vbExclamation
'        End
'     End If
     
'lo tengo que poner cuando vaya a hacer el listado de diferencias shifru.
'abriendo la conexion de multibase
'       AbrirConexionMultibase "mb", "mb"
     
       Load frmIdentifica
       CadenaDesdeOtroForm = ""
       
       'Necesitaremos el archivo arifon.dat
       frmIdentifica.Show vbModal
        
       If CadenaDesdeOtroForm = "" Then
            'NO se ha identificado
            Set Conn = Nothing
            End
       End If
       
       CadenaDesdeOtroForm = ""
       frmLogin.Show vbModal
       If CadenaDesdeOtroForm = "" Then
            'No ha seleccionado nonguna empresa
            Set Conn = Nothing
            End
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass

'        LeerEmpresa 'Carga los Datos de la empresa
        'Carga los Datos B�sicos de la empresa
        LeerDatosEmpresa


        'Cerramos la conexion
        Conn.Close

        'Abre la conexi�n a BDatos:Ariges
        If AbrirConexion() = False Then
            MsgBox "La aplicaci�n no puede continuar sin acceso a los datos. ", vbCritical
            End
        Else
            'Carga Parametros Generales y Contables de la empresa
            LeerParametros
        
        
        End If
                
        'Abrir conexi�n a la BDatos de Contabilidad para acceder a
        'Tablas: Cuentas, Tipos IVA
        If vParamAplic.NumeroConta <> 0 Then
            If AbrirConexionConta() = False Then
                MsgBox "La aplicaci�n no puede continuar sin acceso a los datos. ", vbCritical
                End
            End If
        End If
        
        'Carga los Niveles de cuentas de Contabilidad de la empresa y las fechasINICIO y FIN
        If vParamAplic.NumeroConta <> 0 Then LeerNivelesEmpresa

'        'Gestionar el nombre del PC para la asignacion de PC en el entorno de red
'[Monica]08/10/2009 lo he descomentado pq estaba comentado
        GestionaPC

'        'Otras acciones
        OtrasAcciones
        Screen.MousePointer = vbHourglass
        Load MDIppal
        Screen.MousePointer = vbDefault
        MDIppal.Show
     
     
'     'obric la conexio
'    If AbrirConexionAriagro("root", "aritel") = False Then
'        MsgBox "La aplicaci�n no puede continuar sin acceso a los datos. ", vbCritical
'        End
'    End If
'
'    Load frmIdentifica
'    'CadenaDesdeOtroForm = ""
'
'    'Necesitaremos el archivo login.dat
'    frmIdentifica.Show
    
End Sub

Public Function ComprovaVersio() As Boolean
  
'    Dim RS2 As Recordset
'    Dim cad2 As String
'    Dim major_ul As Integer
'    Dim minor_ul As Integer
'    Dim revis_ul As Integer
'
'    ComprovaVersio = False
'
'    cad2 = "SELECT * FROM ulversio"
'
'    Set RS2 = New ADODB.Recordset
'    RS2.Open cad2, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    If Not RS2.EOF Then
'        major_ul = RS2.Fields!major_ul
'        minor_ul = RS2.Fields!minor_ul
'        revis_ul = RS2.Fields!revis_ul
'    Else
'        MsgBox "Error al consultar la �ltima versi�n disponible", vbCritical
'        'ulVersio = False
'        Exit Function
'    End If
'
'    RS2.Close
'    Set RS2 = Nothing
'
'    If (App.Major <> major_ul) Or (App.Minor <> minor_ul) Or (App.Revision <> revis_ul) Then
'        ComprovaVersio = True
'    End If
'
'    Exit Function
    
End Function

'espera els segon que li digam
Public Function espera(Segundos As Single)
    Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function



Public Function AbrirConexionConta() As Boolean
'Abre

Dim cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionConta = False
    Set ConnConta = Nothing
    Set ConnConta = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnConta.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
                        
'[Monica]23/09/2014: dejamos la conexion igual que en todas las aplicaciones, como en recoleccion
    If vParamAplic.ServidorConta = "" Then vParamAplic.ServidorConta = vConfig.SERVER
                       
    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamAplic.NumeroConta & ";SERVER=" & vParamAplic.ServidorConta & ";"
    cad = cad & ";UID=" & vParamAplic.UsuarioConta
    cad = cad & ";PWD=" & vParamAplic.PasswordConta
    '---- Laura: 29/09/2006
    cad = cad & ";PORT=3306;OPTION=3;STMT=;"
    '----
    '++monica: tema de vista
    cad = cad & "Persist Security Info=true"

    
'    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta" & vParamAplic.NumeroConta & ";SERVER=" & vParamAplic.ServidorConta & ";"
'    cad = cad & ";UID=" & vParamAplic.UsuarioConta
'    cad = cad & ";PASSWORD=" & vParamAplic.PasswordConta
'    cad = cad & ";PORT=3306;OPTION=3;STMT=;"
'    cad = cad & ";Persist Security Info=true"
    
    ConnConta.ConnectionString = cad
    ConnConta.Open
    ConnConta.Execute "Set AUTOCOMMIT = 1"
    AbrirConexionConta = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n contabilidad.", Err.Description
End Function


Public Function AbrirConexionAridoc(Usuario As String, Pass As String) As Boolean
'Abre
Dim cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionAridoc = False
    Set ConnAridoc = Nothing
    Set ConnAridoc = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnAridoc.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
'    cad = "DSN=Aridoc;DESC=MySQL ODBC 3.51 Driver DSN;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
    
    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=aridoc;SERVER=" & vConfig.SERVER & ";"
    cad = cad & ";UID=" & Usuario
    cad = cad & ";PWD=" & Pass
                     
    '++monica:tema del vista
    cad = cad & ";Persist Security Info=true"
    '++
                     
                     
    ConnAridoc.ConnectionString = cad
    ConnAridoc.Open
    ConnAridoc.Execute "Set AUTOCOMMIT = 1"
    AbrirConexionAridoc = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n Aridoc.", Err.Description
End Function

Public Function AbrirConexionMultibase(Usuario As String, Pass As String) As Boolean
Dim cad As String
On Error GoTo EAbrirConexion
    
    AbrirConexionMultibase = False
    
    
    Set ConnMB = Nothing
    Set ConnMB = New Connection
    ConnMB.CursorLocation = adUseClient

'    cad = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=mAgroMB;UID=" & Usuario & ";PWD=" & Pass & ";"""
  '  Debug.Print conn.Version
    cad = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=vAriagroMB;UID=" & Usuario & ";PWD=" & Pass & ";"
    cad = cad & ";Persist Security Info=true"

    ConnMB.ConnectionString = cad
    ConnMB.Open
    AbrirConexionMultibase = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n Multibase.", Err.Description
End Function


'Public Function AbrirConexionAuxCon(Empresa As String, Usuario As String, Pass As String) As Boolean
'Dim cad As String
'Dim nomConta As String 'nombre de la BD de la contabilidad
'Dim serConta As String 'servidor donde esta la BD de la contabilidad
'On Error GoTo EAbrirConexion
'
'    AbrirConexionAuxCon = False
'
'    Set ConnAuxCon = Nothing
'    Set ConnAuxCon = New Connection
''    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
'    ConnAuxCon.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
'
'    'Obtener la BD de contabilidad
''    SQL = "select bdaconta FROM paramcon WHERE codempre=" & codEmpre
'    serConta = "serconta"
'    nomConta = DevuelveDesdeBDNew(2, "sparam", "bdaconta", "codempre", Empresa, "N", serConta)
''    vEmpresa.BDConta = nomConta
'    If nomConta <> "" Then
'    '    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamConta.NumeroConta & ";SERVER=" & vParamConta.ServidorConta & ";"
'    '    cad = cad & ";UID=" & vParamConta.UsuarioConta
'    '    cad = cad & ";PWD=" & vParamConta.PasswordConta
'    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        If serConta <> "" Then 'especificamos servidor
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & serConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        Else 'por defecto cogera la BD del servidor que haya en el ODBC
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        End If
'        ConnAuxCon.ConnectionString = cad
'        ConnAuxCon.Open
'        ConnAuxCon.Execute "Set AUTOCOMMIT = 1"
'        AbrirConexionAuxCon = True
'    Else
'        AbrirConexionAuxCon = False
'    End If
'    Exit Function
'EAbrirConexion:
'    MuestraError Err.Number, "Abrir conexi�n Contabilidad.", Err.Description
'End Function

Public Function CerrarConexionConta()
  'Cerramos la conexion con BD: Contabilidad
  On Error Resume Next
   ConnConta.Close
   If Err.Number <> 0 Then Err.Clear
End Function

Public Function CerrarConexionUsuarios()
  'Cerramos la conexion con BD: Usuarios
  On Error Resume Next
   Conn.Close
   If Err.Number <> 0 Then Err.Clear
End Function

Public Function CerrarConexionAridoc()
  'Cerramos la conexion con BD: Aridoc
  On Error Resume Next
   ConnAridoc.Close
   If Err.Number <> 0 Then Err.Clear
End Function

Public Function CerrarConexionMultibase()
  'Cerramos la conexion con BD: Multibase
  On Error Resume Next
   ConnMB.Close
   If Err.Number <> 0 Then Err.Clear
End Function


Public Function LeerDatosEmpresa()
 'Crea instancia de la clase Cempresa con los valores en
 'Tabla: ArigesEmpresa
 'BDatos: Usuarios
 
        Set vEmpresa = New Cempresa
        If vEmpresa.LeerDatos = 1 Then
            MsgBox "No se han podido cargar datos empresa (BD:usuarios). Debe configurar la aplicaci�n.", vbExclamation
            Set vEmpresa = Nothing
        End If
            
End Function


Public Sub LeerParametros()
'Crea instancia de la clase Cempresa con los valores en
'Tabla: Empresas
'BDatos: PTours y Conta
 Dim devuelve As String
 
    'Parametros Generales
    Set vParam = New Cparametros
    If vParam.leer() = 1 Then
        devuelve = "No se han podido cargar los Par�metros Generales.(empresas)" & vbCrLf
        MsgBox devuelve & " Debe configurar la aplicaci�n.", vbExclamation
        Set vParam = Nothing
    End If
    
    ' ### [Monica] 06/09/2006
    ' a�adido
    Set vParamAplic = New CParamAplic
    If vParamAplic.leer = 1 Then
        MsgBox "No se han podido cargar los Par�metros de la Aplicaci�n(sparam). Debe configurar la aplicaci�n.", vbExclamation

        Set vParamAplic = Nothing
        Exit Sub
'    Else
'    If vParamAplic.NumeroConta <> 0 Then
'
'        'Abrir conexi�n a la BDatos de Contabilidad para acceder a
'        'Tablas: Cuentas, Tipos IVA,...
'        If AbrirConexionConta(vParamAplic.UsuarioConta, vParamAplic.PasswordConta) = False Then
'            MsgBox "La aplicaci�n no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
'            AccionesCerrar
'            End
'        End If
'    ' ### [Monica] 06/09/2006
'    ' comento los niveles de contabilidad pq solo tengo las cuentas
'        If vEmpresa.LeerNiveles() = False Then  'De Contabilidad
'            MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicaci�n.", vbExclamation
'            AccionesCerrar
'            End
'        End If
'
'        FechasEjercicioConta FIni, FFin
'
'    End If
    End If

'    Set vParam = New Cparametros
'    If vParam.Leer = False Then   'De AriGasol
'        MsgBox "No se han podido cargar los par�metros de la empresa. Debe configurar la aplicaci�n.", vbExclamation
'        Set vEmpresa = Nothing
'        Set vSesion = Nothing
'        Set Conn = Nothing
'        End
'    End If
End Sub


Public Function PonerDatosPpal()
'    If Not vEmpresa Is Nothing Then
'        MDIppal.Caption = "AriAgro" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  Empresa: " & vEmpresa.nomEmpre
'    End If
    
    If vParam Is Nothing Then
        MDIppal.Caption = "AriAgro - Gesti�n Comercial" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & " FALTA CONFIGURAR"
    Else
        MDIppal.Caption = "AriAgro - Gesti�n Comercial" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  Empresa: " & vParam.NombreEmpresa & _
                  " - Campa�a: " & vParam.FecIniCam & " - " & vParam.FecFinCam & "   -  Usuario: " & vUsu.Nombre
    End If

    
    If Err.Number <> 0 Then MuestraError Err.Description, "Poniendo datos de la pantalla principal", Err.Description
End Function

    

Public Sub MuestraError(numero As Long, Optional cadena As String, Optional Desc As String)
    Dim cad As String
    Dim Aux As String
    
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    cad = "Se ha producido un error: " & vbCrLf
    If cadena <> "" Then
        cad = cad & vbCrLf & cadena & vbCrLf & vbCrLf
    End If
    'Numeros de errores que contolamos
    If Conn.Errors.Count > 0 Then
        ControlamosError Aux
        Conn.Errors.Clear
    Else
        Aux = ""
    End If
    If Aux <> "" Then Desc = Aux
    If Desc <> "" Then cad = cad & vbCrLf & Desc & vbCrLf & vbCrLf
    If Aux = "" Then cad = cad & "N�mero: " & numero & vbCrLf & "Descripci�n: " & Error(numero)
    MsgBox cad, vbExclamation
End Sub

Public Function DBSet(vData As Variant, tipo As String, Optional EsNulo As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
Dim cad As String

        If IsNull(vData) Then
            DBSet = ValorNulo
            Exit Function
        End If

        If tipo <> "" Then
            Select Case tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If EsNulo = "N" Then
                            DBSet = "''"
                        Else
                            DBSet = ValorNulo
                        End If
                    Else
                        cad = (CStr(vData))
                        NombreSQL cad
                        DBSet = "'" & cad & "'"
                    End If
                    
                Case "N"    'Numero
                    If vData = "" Or vData = 0 Then
                        If EsNulo <> "" Then
                            If EsNulo = "S" Then
                                DBSet = ValorNulo
                            Else
                                DBSet = 0
                            End If
                        Else
                            DBSet = 0
                        End If
                    Else
                        cad = CStr(ImporteFormateado(CStr(vData)))
                        DBSet = TransformaComasPuntos(cad)
                    End If
                    
                Case "F"    'Fecha
'                     '==David
''                    DBLet = "0:00:00"
'                     '==Laura
                    If vData = "" Then
                        If EsNulo = "S" Then
                            DBSet = ValorNulo
                        Else
                            DBSet = "'1900-01-01'"
                        End If
                    Else
                        DBSet = "'" & Format(vData, FormatoFecha) & "'"
                    End If
                    
                Case "FH" 'Fecha/Hora
                    If vData = "" Then
                        If EsNulo = "S" Then DBSet = ValorNulo
                    Else
                        DBSet = "'" & Format(vData, "yyyy-mm-dd hh:mm:ss") & "'"
                    End If
                    
                Case "H" 'Hora
                    If vData = "" Then
                    Else
                        DBSet = "'" & Format(vData, "hh:mm:ss") & "'"
                    End If
                    
                Case "B"  'Boolean
                    If vData Then
                        DBSet = 1
                    Else
                        DBSet = 0
                    End If
            End Select
        End If
End Function

Public Function DBLetMemo(vData As Variant) As Variant
    On Error Resume Next
    
    DBLetMemo = vData
    
    
    
    If Err.Number <> 0 Then
        Err.Clear
        DBLetMemo = ""
    End If
End Function



Public Function DBLet(vData As Variant, Optional tipo As String) As Variant
'Para cuando recupera Datos de la BD
    If IsNull(vData) Then
        DBLet = ""
        If tipo <> "" Then
            Select Case tipo
                Case "T"    'Texto
                    DBLet = ""
                Case "N"    'Numero
                    DBLet = 0
                Case "F"    'Fecha
                     '==David
'                    DBLet = "0:00:00"
                     '==Laura
'                     DBLet = "0000-00-00"
                      DBLet = ""
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo antes de bloquear
'   Prepara la conexion para bloquear
Public Sub PreparaBloquear()
    Conn.Execute "commit"
    Conn.Execute "set autocommit=0"
End Sub

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo despues de un bloque
'   Prepara la conexion para bloquear
Public Sub TerminaBloquear()
    Conn.Execute "commit"
    Conn.Execute "set autocommit=1"
End Sub

'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256,98
'   Tiene que venir num�rico
Public Function ImporteFormateado(Importe As String) As Currency
Dim i As Integer

    If Importe = "" Then
        ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            i = InStr(1, Importe, ".")
            If i > 0 Then Importe = Mid(Importe, 1, i - 1) & Mid(Importe, i + 1)
        Loop Until i = 0
        ImporteFormateado = Importe
    End If
End Function

' ### [Monica] 11/09/2006
Public Function ImporteSinFormato(cadena As String) As String
Dim i As Integer
'Quitamos puntos
Do
    i = InStr(1, cadena, ".")
    If i > 0 Then cadena = Mid(cadena, 1, i - 1) & Mid(cadena, i + 1)
Loop Until i = 0
ImporteSinFormato = TransformaPuntosComas(cadena)
End Function



'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(cadena As String) As String
Dim i As Integer
    Do
        i = InStr(1, cadena, ",")
        If i > 0 Then
            cadena = Mid(cadena, 1, i - 1) & "." & Mid(cadena, i + 1)
        End If
    Loop Until i = 0
    TransformaComasPuntos = cadena
End Function

'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef cadena As String)
Dim J As Integer
Dim i As Integer
Dim Aux As String
    J = 1
    Do
        i = InStr(J, cadena, "'")
        If i > 0 Then
            Aux = Mid(cadena, 1, i - 1) & "\"
            cadena = Aux & Mid(cadena, i)
            J = i + 2
        End If
    Loop Until i = 0
End Sub

Public Function EsFechaOKString(ByRef T As String) As Boolean
Dim cad As String
    
    cad = T
    If InStr(1, cad, "/") = 0 Then
        If Len(T) = 8 Then
            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        Else
            If Len(T) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        End If
    End If
    If IsDate(cad) Then
        EsFechaOKString = True
        T = Format(cad, "dd/mm/yyyy")
    Else
        EsFechaOKString = False
    End If
End Function

Public Function DevNombreSQL(cadena As String) As String
Dim J As Integer
Dim i As Integer
Dim Aux As String
    J = 1
    Do
        i = InStr(J, cadena, "'")
        If i > 0 Then
            Aux = Mid(cadena, 1, i - 1) & "\"
            cadena = Aux & Mid(cadena, i)
            J = i + 2
        End If
    Loop Until i = 0
    DevNombreSQL = cadena
End Function


Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional tipo As String, Optional ByRef otroCampo As String) As String
    Dim Rs As Recordset
    Dim cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    cad = "Select " & kCampo
    If otroCampo <> "" Then cad = cad & ", " & otroCampo
    cad = cad & " FROM " & Ktabla
    cad = cad & " WHERE " & Kcodigo & " = "
    If tipo = "" Then tipo = "N"
    Select Case tipo
    Case "N"
        'No hacemos nada
        cad = cad & ValorCodigo
    Case "T", "F"
        cad = cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
    
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        DevuelveDesdeBD = DBLet(Rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EDevuelveDesdeBD:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function



''Este metodo sustituye a DevuelveDesdeBD
''Funciona para claves primarias formadas por 2 campos
'Public Function DevuelveDesdeBDnew(vBD As Byte, Ktabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef otroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String) As String
''IN: vBD --> Base de Datos a la que se accede
'Dim RS As Recordset
'Dim cad As String
'Dim Aux As String
'
'On Error GoTo EDevuelveDesdeBDnew
'    DevuelveDesdeBDnew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
'    cad = "Select " & kCampo
'    If otroCampo <> "" Then cad = cad & ", " & otroCampo
'    cad = cad & " FROM " & Ktabla
'    cad = cad & " WHERE " & Kcodigo1 & " = "
'    If tipo1 = "" Then tipo1 = "N"
'    Select Case tipo1
'        Case "N"
'            'No hacemos nada
'            If IsNumeric(valorCodigo1) Then
'                cad = cad & Val(valorCodigo1)
'            Else
'                MsgBox "El campo debe ser num�rico.", vbExclamation
'                DevuelveDesdeBDnew = "Error"
'                Exit Function
'            End If
'        Case "T", "F"
'            cad = cad & "'" & valorCodigo1 & "'"
'        Case Else
'            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
'            Exit Function
'    End Select
'
'    If KCodigo2 <> "" Then
'        cad = cad & " AND " & KCodigo2 & " = "
'        If tipo2 = "" Then tipo2 = "N"
'        Select Case tipo2
'        Case "N"
'            'No hacemos nada
'            If ValorCodigo2 = "" Then
'                cad = cad & "-1"
'            Else
'                cad = cad & Val(ValorCodigo2)
'            End If
'        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
'        Case "F"
'            cad = cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
'        Case Else
'            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
'            Exit Function
'        End Select
'    End If
'
'
'    'Creamos el sql
'    Set RS = New ADODB.Recordset
'
'    Select Case vBD
'        Case cAgro 'vBD=1: PlannerTours
'            RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'        Case cConta 'BD 2: Contabilidad
'            RS.Open cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
'        Case 3 'vBD=3: contabilidad distinta a la de la empresa conectada
'            RS.Open cad, ConnAuxCon, adOpenForwardOnly, adLockOptimistic, adCmdText
'    End Select
''    If vBD = cAgro Then 'vBD=1: PlannerTours
''        RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
''    ElseIf vBD = cConta Then  'BD 2: Contabilidad
''        RS.Open cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
''    End If
'
'    If Not RS.EOF Then
'        DevuelveDesdeBDnew = DBLet(RS.Fields(0))
'        If otroCampo <> "" Then otroCampo = DBLet(RS.Fields(1))
'    End If
'    RS.Close
'    Set RS = Nothing
'    Exit Function
'
'EDevuelveDesdeBDnew:
'        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
'End Function


'LAURA
'Este metodo sustituye a DevuelveDesdeBD
'Funciona para claves primarias formadas por 3 campos
Public Function DevuelveDesdeBDNew(vBD As Byte, Ktabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef otroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String, Optional KCodigo3 As String, Optional ValorCodigo3 As String, Optional tipo3 As String) As String
'IN: vBD --> Base de Datos a la que se accede
Dim Rs As Recordset
Dim cad As String
Dim Aux As String
    
On Error GoTo EDevuelveDesdeBDnew
    DevuelveDesdeBDNew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
    cad = "Select " & kCampo
    If otroCampo <> "" Then cad = cad & ", " & otroCampo
    cad = cad & " FROM " & Ktabla
    If Kcodigo1 <> "" Then
        cad = cad & " WHERE " & Kcodigo1 & " = "
        If tipo1 = "" Then tipo1 = "N"
    Select Case tipo1
        Case "N"
            'No hacemos nada
            cad = cad & Val(valorCodigo1)
        Case "T"
            cad = cad & DBSet(valorCodigo1, "T")
        Case "F"
            cad = cad & DBSet(valorCodigo1, "F")
        Case Else
            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
            Exit Function
    End Select
    End If
    
    If KCodigo2 <> "" Then
        cad = cad & " AND " & KCodigo2 & " = "
        If tipo2 = "" Then tipo2 = "N"
        Select Case tipo2
        Case "N"
            'No hacemos nada
            If ValorCodigo2 = "" Then
                cad = cad & "-1"
            Else
                cad = cad & Val(ValorCodigo2)
            End If
        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
            cad = cad & DBSet(ValorCodigo2, "T")
        Case "F"
            cad = cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    If KCodigo3 <> "" Then
        cad = cad & " AND " & KCodigo3 & " = "
        If tipo3 = "" Then tipo3 = "N"
        Select Case tipo3
        Case "N"
            'No hacemos nada
            If ValorCodigo3 = "" Then
                cad = cad & "-1"
            Else
                cad = cad & Val(ValorCodigo3)
            End If
        Case "T"
            cad = cad & "'" & ValorCodigo3 & "'"
        Case "F"
            cad = cad & "'" & Format(ValorCodigo3, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo3 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    
    Select Case vBD
        Case cAgro 'BD 1: Ariges
            Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Case cConta 'BD 2: Conta
            Rs.Open cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
        Case cAridoc 'BD 2: Aridoc
            Rs.Open cad, ConnAridoc, adOpenForwardOnly, adLockOptimistic, adCmdText
    End Select
    
    If Not Rs.EOF Then
        DevuelveDesdeBDNew = DBLet(Rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
EDevuelveDesdeBDnew:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function




'CESAR
Public Function DevuelveDesdeBDnew2(kBD As Integer, kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional tipo As String, Optional num As Byte, Optional ByRef otroCampo As String) As String
Dim Rs As Recordset
Dim cad As String
Dim Aux As String
Dim v_aux As Integer
Dim campo As String
Dim Valor As String
Dim tip As String

On Error GoTo EDevuelveDesdeBDnew2
DevuelveDesdeBDnew2 = ""

cad = "Select " & kCampo
If otroCampo <> "" Then cad = cad & ", " & otroCampo
cad = cad & " FROM " & Ktabla

If Kcodigo <> "" Then cad = cad & " where "

For v_aux = 1 To num
    campo = RecuperaValor(Kcodigo, v_aux)
    Valor = RecuperaValor(ValorCodigo, v_aux)
    tip = RecuperaValor(tipo, v_aux)
        
    cad = cad & campo & "="
    If tip = "" Then tipo = "N"
    
    Select Case tip
            Case "N"
                'No hacemos nada
                cad = cad & Valor
            Case "T", "F"
                cad = cad & "'" & Valor & "'"
            Case Else
                MsgBox "Tipo : " & tip & " no definido", vbExclamation
            Exit Function
    End Select
    
    If v_aux < num Then cad = cad & " AND "
  Next v_aux

'Creamos el sql
Set Rs = New ADODB.Recordset
Select Case kBD
    Case 1
        Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
End Select

If Not Rs.EOF Then
    DevuelveDesdeBDnew2 = DBLet(Rs.Fields(0))
    If otroCampo <> "" Then otroCampo = DBLet(Rs.Fields(1))
Else
     If otroCampo <> "" Then otroCampo = ""
End If
Rs.Close
Set Rs = Nothing
Exit Function
EDevuelveDesdeBDnew2:
    MuestraError Err.Number, "Devuelve DesdeBDnew2.", Err.Description
End Function


Public Function EsEntero(Texto As String) As Boolean
Dim i As Integer
Dim C As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEntero = False

    If Not IsNumeric(Texto) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            i = InStr(L, Texto, ".")
            If i > 0 Then
                L = i + 1
                C = C + 1
            End If
        Loop Until i = 0
        If C > 1 Then res = False
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                i = InStr(L, Texto, ",")
                If i > 0 Then
                    L = i + 1
                    C = C + 1
                End If
            Loop Until i = 0
            If C > 1 Then res = False
        End If
        
    End If
        EsEntero = res
End Function

Public Function TransformaPuntosComas(cadena As String) As String
    Dim i As Integer
    Do
        i = InStr(1, cadena, ".")
        If i > 0 Then
            cadena = Mid(cadena, 1, i - 1) & "," & Mid(cadena, i + 1)
        End If
        Loop Until i = 0
    TransformaPuntosComas = cadena
End Function

Public Sub InicializarFormatos()
    FormatoFecha = "yyyy-mm-dd"
    FormatoHora = "hh:mm:ss"
'    FormatoFechaHora = "yyyy-mm-dd hh:mm:ss"
    FormatoImporte = "#,###,###,##0.00"  'Decimal(12,2)
    FormatoPrecio = "##,##0.000"  'Decimal(8,3) antes decimal(10,4)
'    FormatoCantidad = "##,###,##0.00"   'Decimal(10,2)
    FormatoPorcen = "##0.00" 'Decima(5,2) para porcentajes
    
    FormatoDec10d2 = "##,###,##0.00"   'Decimal(10,2)
    FormatoDec10d3 = "##,###,##0.000"   'Decimal(10,3)
    FormatoDec5d4 = "0.0000"   'Decimal(5,4)
    FormatoDec8d4 = "###0.0000" ' Decimal(8,4)
    FormatoDec8d2 = "###,##0.00" ' Decimal(8,2)
    FormatoDec6d4 = "#0.0000" ' Decimal(6,4)
    FormatoExp = "0000000000"
'    FormatoKms = "#,##0.00##" 'Decimal(8,4)
End Sub


Public Sub AccionesCerrar()
'cosas que se deben hacen cuando finaliza la aplicacion
    On Error Resume Next
    
    'cerrar clases q estan abiertas durante la ejecucion
    Set vEmpresa = Nothing
    Set vSesion = Nothing
    
'    Set vParam = Nothing
'    Set vParamAplic = Nothing
'    Set vParamConta = Nothing
    
    
    'Cerrar Conexiones a bases de datos
    Conn.Close
    ConnConta.Close
    Set Conn = Nothing
    Set ConnConta = Nothing
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function OtrosPCsContraAplicacion() As String
Dim MiRS As Recordset
Dim cad As String
Dim Equipo As String

    Set MiRS = New ADODB.Recordset
    cad = "show processlist"
    MiRS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    While Not MiRS.EOF
        If MiRS.Fields(3) = vUsu.CadenaConexion Then
            Equipo = MiRS.Fields(2)
            'Primero quitamos los dos puntos del puerot
            NumRegElim = InStr(1, Equipo, ":")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            'El punto del dominio
            NumRegElim = InStr(1, Equipo, ".")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            Equipo = UCase(Equipo)
            
            If Equipo <> vUsu.PC Then
                    If Equipo <> "LOCALHOST" Then
                        If InStr(1, cad, Equipo & "|") = 0 Then cad = cad & Equipo & "|"
                    End If
            End If
        End If
        'Siguiente
        MiRS.MoveNext
    Wend
    NumRegElim = 0
    MiRS.Close
    Set MiRS = Nothing
    OtrosPCsContraAplicacion = cad
End Function


Public Function UsuariosConectados() As Boolean
Dim i As Integer
Dim cad As String
Dim metag As String
Dim Sql As String
cad = OtrosPCsContraAplicacion
UsuariosConectados = False
If cad <> "" Then
    UsuariosConectados = True
    i = 1
    metag = "Los siguientes PC's est�n conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
    Do
        Sql = RecuperaValor(cad, i)
        If Sql <> "" Then
            metag = metag & "    - " & Sql & vbCrLf
            i = i + 1
        End If
    Loop Until Sql = ""
    MsgBox metag, vbExclamation
End If
End Function

'Usuario As String, Pass As String --> Directamente el usuario
Public Function AbrirConexion() As Boolean
Dim cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexion = False
    Set Conn = Nothing
    Set Conn = New Connection
    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    Conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
'[Monica]23/09/2014: dejamos la conexion igual que en todas las aplicaciones, como en recoleccion
    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & vUsu.CadenaConexion & ";SERVER=" & vConfig.SERVER & ";"
    cad = cad & ";UID=" & vConfig.User
    cad = cad & ";PWD=" & vConfig.password
    cad = cad & ";Persist Security Info=true"
    
'    cad = "DSN=vAriagro;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & vUsu.CadenaConexion & ";UID=" & vConfig.User & ";PASSWORD=" & vConfig.password & ";PORT=3306;OPTION=3;STMT=;"
'    cad = cad & ";Persist Security Info=true"

    Conn.ConnectionString = cad
    Conn.Open
    Conn.Execute "Set AUTOCOMMIT = 1"
    AbrirConexion = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n.", Err.Description
End Function

Public Function LeerEmpresaParametros()
        'Abrimos la empresa
        Set vEmpresa = New Cempresa
        If vEmpresa.LeerDatos = 1 Then
            MsgBox "No se han podido cargar datos empresa. Debe configurar la aplicaci�n.", vbExclamation
            Set vEmpresa = Nothing
        End If
            
            
'        Set vParamAplic = New CParamAplic
'        If vParamAplic.Leer() = 1 Then
'            MsgBox "No se han podido cargar los par�metros. Debe configurar la aplicaci�n.", vbExclamation
'            Set vParamAplic = Nothing
'        End If
        
        If Not (vEmpresa Is Nothing) Then 'And Not (vParamAplic Is Nothing) Then

            CadenaDesdeOtroForm = ""
        End If
        
        
End Function

Private Sub GestionaPC()
CadenaDesdeOtroForm = ComputerName
If CadenaDesdeOtroForm <> "" Then
    FormatoFecha = DevuelveDesdeBD("codpc", "usuarios.pcs", "nompc", CadenaDesdeOtroForm, "T")
    If FormatoFecha = "" Then
        NumRegElim = 0
        FormatoFecha = "Select max(codpc) from usuarios.pcs"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open FormatoFecha, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            NumRegElim = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        NumRegElim = NumRegElim + 1
        If NumRegElim > 32000 Then
            MsgBox "Error en numero de PC's activos. Demasiados PC en BD. Llame a soporte t�cnico.", vbCritical
            End
        End If
        FormatoFecha = "INSERT INTO usuarios.pcs (codpc, nompc) VALUES (" & NumRegElim & ", '" & CadenaDesdeOtroForm & "')"
        Conn.Execute FormatoFecha
    End If
End If
End Sub


Private Sub OtrasAcciones()
On Error Resume Next

    FormatoFecha = "yyyy-mm-dd"
    FormatoFechaHora = "yyyy-mm-dd hh:mm:ss"
    FormatoImporte = "#,###,###,##0.00"
    FormatoCantidad = "##,###,##0.00"   'Decimal(10,2)
    FormatoDescuento = "#0.00" 'Decima(4,2)

    InicializarFormatos

    'Borramos uno de los archivos temporales
    If Dir(App.path & "\ErrActua.txt") <> "" Then Kill App.path & "\ErrActua.txt"
    
    
    'Borramos tmp bloqueos
    'Borramos temporal
    CadenaDesdeOtroForm = OtrosPCsContraContabiliad
    NumRegElim = Len(CadenaDesdeOtroForm)
    If NumRegElim = 0 Then
        CadenaDesdeOtroForm = ""
    Else
        CadenaDesdeOtroForm = " WHERE codusu = " & vUsu.Codigo
    End If
    Conn.Execute "Delete from zBloqueos " & CadenaDesdeOtroForm
    CadenaDesdeOtroForm = ""
    NumRegElim = 0
    
    
End Sub

Public Function OtrosPCsContraContabiliad() As String
Dim MiRS As Recordset
Dim cad As String
Dim Equipo As String

    Set MiRS = New ADODB.Recordset
    cad = "show processlist"
    MiRS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    While Not MiRS.EOF
        If MiRS.Fields(3) = vUsu.CadenaConexion Then
            Equipo = MiRS.Fields(2)
            'Primero quitamos los dos puntos del puerot
            NumRegElim = InStr(1, Equipo, ":")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            'El punto del dominio
            NumRegElim = InStr(1, Equipo, ".")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            Equipo = UCase(Equipo)
            
            If Equipo <> vUsu.PC Then
                    If Equipo <> "LOCALHOST" Then
                        If InStr(1, cad, Equipo & "|") = 0 Then cad = cad & Equipo & "|"
                    End If
            End If
        End If
        'Siguiente
        MiRS.MoveNext
    Wend
    NumRegElim = 0
    MiRS.Close
    Set MiRS = Nothing
    OtrosPCsContraContabiliad = cad
End Function

Public Function ComprobarEmpresaBloqueada(Codusu As Long, ByRef Empresa As String) As Boolean
Dim cad As String

ComprobarEmpresaBloqueada = False

'Antes de nada, borramos las entradas de usuario, por si hubiera kedado algo
Conn.Execute "Delete from usuarios.vbloqbd where codusu=" & Codusu

'Ahora comprobamos k nadie bloquea la BD
cad = DevuelveDesdeBD("codusu", "usuarios.vbloqbd", "conta", Empresa, "T")
If cad <> "" Then
    'En teoria esta bloqueada. Puedo comprobar k no se haya kedado el bloqueo a medias
    
    Set miRsAux = New ADODB.Recordset
    cad = "show processlist"
    miRsAux.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    While Not miRsAux.EOF
        If miRsAux.Fields(3) = Empresa Then
            cad = miRsAux.Fields(2)
            miRsAux.MoveLast
        End If
    
        'Siguiente
        miRsAux.MoveNext
    Wend
    
    If cad = "" Then
        'Nadie esta utilizando la aplicacion, luego se puede borrar la tabla
        Conn.Execute "Delete from usuarios.vbloqbd where conta ='" & Empresa & "'"
        
    Else
        MsgBox "BD bloqueada.", vbCritical
        ComprobarEmpresaBloqueada = True
    End If
End If

Conn.Execute "commit"
End Function

Public Function AbrirConexionUsuarios() As Boolean
Dim cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionUsuarios = False
    Set Conn = Nothing
    Set Conn = New Connection
    'Conn.CursorLocation = adUseClient
    Conn.CursorLocation = adUseServer
    
'[Monica]23/09/2014: dejamos la conexion igual que en todas las aplicaciones, como en recoleccion
    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=usuarios;SERVER=" & vConfig.SERVER
    cad = cad & ";UID=" & vConfig.User
    cad = cad & ";PWD=" & vConfig.password
    '++monica: tema del vista
    cad = cad & ";Persist Security Info=true"
'    '++
    
    
'    cad = "DSN=vUsuarios;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=usuarios;"
'    cad = cad & "SERVER=" & vConfig.SERVER & ";UID=" & vConfig.User & ";PASSWORD=" & vConfig.password & ";PORT=3306;OPTION=3;STMT=;"
'    cad = cad & ";Persist Security Info=true"

    
    Conn.ConnectionString = cad
    Conn.Open
    AbrirConexionUsuarios = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n usuarios.", Err.Description
End Function

Public Sub CommitConexion()
On Error Resume Next
    Conn.Execute "Commit"
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function LeerNivelesEmpresa()
 'Crea instancia de la clase Cempresa con los valores en
 'Tabla: Empresa
 'BDatos: Conta
        
        If vEmpresa.LeerNiveles = 1 Then
            MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicaci�n.", vbExclamation
            Set vEmpresa = Nothing
        End If
            
End Function

'--------------------------------------------------------------------
'-------------------------------------------------------------------
'Para el envio de los mails
Public Function PrepararCarpetasEnvioMail(Optional NoBorrar As Boolean) As Boolean
    On Error GoTo EPrepararCarpetasEnvioMail
    PrepararCarpetasEnvioMail = False

    If Dir(App.path & "\temp", vbDirectory) = "" Then
        MkDir App.path & "\temp"
    Else
        If Not NoBorrar Then
            If Dir(App.path & "\temp\*.*", vbArchive) <> "" Then Kill App.path & "\temp\*.*"
        End If
    End If


    PrepararCarpetasEnvioMail = True
    Exit Function
EPrepararCarpetasEnvioMail:
    MuestraError Err.Number, "", "Preparar Carpetas Envio Mail "
End Function


'------------------------------------------------------------------
'   Comprobara si una daterminada fecha esta o no en los ejercicios
'   contables (actual y siguiente)
'   Dando un O: SI. Correcto. Ok
'            1: Inferior
'            2: Superior

Public Function EsFechaOKConta(Fecha As Date) As Byte
Dim F2 As Date

    If vEmpresa.FechaIni > Fecha Then
        EsFechaOKConta = 1
    Else
        F2 = DateAdd("yyyy", 1, vEmpresa.FechaFin)
        If Fecha > F2 Then
            EsFechaOKConta = 2
        Else
            'OK. Dentro de los ejercicios contables
            EsFechaOKConta = 0
        End If
    End If

End Function

