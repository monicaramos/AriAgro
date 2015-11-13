Attribute VB_Name = "LibExport"
Option Explicit
Public dbAriagro As BaseDatos
Public vConfig As Configuracion


Sub Main()

    'Vemos datos de ConfigAgro.ini
    Set vConfig = New Configuracion
    If vConfig.Leer = 1 Then

         MsgBox "MAL CONFIGURADO", vbCritical
         End
         Exit Sub
    End If

    '-- Abrir las bases de datos que usará la utilidad
    Set dbAriagro = New BaseDatos
    'dbAriagro.abrir "vAriagro", "root", "aritel"
    If dbAriagro.abrir_MYSQL(vConfig.SERVER, "ariagro", vConfig.User, vConfig.password) Then
        frmUtExport.Show
    Else
        MsgBox "Error abriendo conexión", vbExclamation
    End If
End Sub

Sub CargarTodosLosCampos()
    '-- Utilidad que carga todos los campos de la base de datos
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim cmp As GRPTC_Campo
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select a.*, b.codprodu from rcampos as a , variedades as b where b.codvarie = a.codvarie"
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            '-- creamos el objeto auxiliar que montará el XML de trazatec
            Set cmp = New GRPTC_Campo
            '-- vamos cargando los diferentes valores
            cmp.codsocio = Rs!codsocio
            cmp.codcampo = Rs!codcampo
            cmp.codprodu = Rs!codprodu
            cmp.codvarie = Rs!codvarie
            cmp.codparti = Rs!codparti
            cmp.hanegada = 0 ' no interesa en trazatec
            cmp.numarbol = 0 ' tampoco interesa
            cmp.poligono = Rs!poligono
            '-- Y ahora el objeto chivato para grabar
            Set chv = New GRPTC_Chivato
            chv.Id = 0 'ya lo montará en el momento de la grabación
            chv.BD_Org = "AGRO"
            '[Monica]16/11/2011: solo en Alzira está SCAMP1
            If Cooperativa = 4 Then
                chv.Tabla = "SCAMP1"
            Else
                chv.Tabla = "SCAMPO"
            End If
            chv.Oper = "I"
            chv.fecha = Format(Date, "dd/mm/yyyy")
            chv.Sep = "&"
            chv.Clv_Old = ""
            chv.Clv_New = CStr(cmp.codcampo)
            chv.XML = cmp.GenXML
            chv.Grabar
            Rs.MoveNext
        Wend
    End If
End Sub

Sub CargarUnCampo(codcampo As Long, Tipo As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim cmp As GRPTC_Campo
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select a.*, b.codprodu from rcampos as a , variedades as b where b.codvarie = a.codvarie"
    Sql = Sql & " and a.codcampo = " & CStr(codcampo)
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set cmp = New GRPTC_Campo
        '-- vamos cargando los diferentes valores
        cmp.codsocio = Rs!codsocio
        cmp.codcampo = Rs!codcampo
        cmp.codprodu = Rs!codprodu
        cmp.codvarie = Rs!codvarie
        cmp.codparti = Rs!codparti
        cmp.hanegada = 0 ' no interesa en trazatec
        cmp.numarbol = 0 ' tampoco interesa
        cmp.poligono = Rs!poligono
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        '[Monica]16/11/2011: solo en Alzira está SCAMP1
        If Cooperativa = 4 Then
            chv.Tabla = "SCAMP1"
        Else
            chv.Tabla = "SCAMPO"
        End If
        chv.Oper = Tipo
        chv.fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
'[Monica] 31/12/2009 solo el campo
'        chv.Clv_New = CStr(cmp.codsocio) & _
'                            "&" & CStr(cmp.codcampo) & _
'                            "&" & CStr(cmp.codprodu) & _
'                            "&" & CStr(cmp.codvarie)
        
        chv.Clv_New = CStr(cmp.codcampo)
        
        
        chv.XML = cmp.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            chv.Clv_Old = Rs!codsocio & "&" & Rs!codcampo & "&" & Rs!codprodu & "&" & Rs!codvarie 'chv.Clv_New
        End If
        chv.Grabar
    End If
End Sub


Sub CargarUnSocio(codsocio As Long, Tipo As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim soc As GRPTC_Socio
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select * from rsocios "
    Sql = Sql & " where codsocio = " & CStr(codsocio)
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set soc = New GRPTC_Socio
        '-- vamos cargando los diferentes valores
        soc.codsocio = Rs!codsocio
        soc.nifsocio = Rs!nifsocio
        soc.nomsocio = Rs!nomsocio
        soc.domsocio = Rs!dirsocio
        soc.telsocio = DBLet(Rs!telsoci1)
        soc.codpobla = 0 ' algo hay que hacer
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        chv.Tabla = "SSOCIO"
        chv.Oper = Tipo
        chv.fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = CStr(soc.codsocio)
        chv.XML = soc.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            chv.Clv_Old = chv.Clv_New
        End If
        chv.Grabar
    End If
End Sub


Sub CargarUnaPoblacion(codpobla As String, Tipo As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim pob As GRPTC_Poblacion
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select * from rpueblos "
    Sql = Sql & " where codpobla = '" & codpobla & "'"
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set pob = New GRPTC_Poblacion
        '-- vamos cargando los diferentes valores
        pob.codpobla = Rs!codpobla
        pob.despobla = Rs!despobla
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        chv.Tabla = "SPOBLA"
        chv.Oper = Tipo
        chv.fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = pob.codpobla
        chv.XML = pob.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            chv.Clv_Old = chv.Clv_New
        End If
        chv.Grabar
    End If
End Sub

Sub CargarUnaCuadrilla(codcapat As Long, Tipo As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim cua As GRPTC_Cuadrilla
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select * from rcapataz "
    Sql = Sql & " where codcapat = " & CStr(codcapat)
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set cua = New GRPTC_Cuadrilla
        '-- vamos cargando los diferentes valores
        cua.codcapat = Rs!codcapat
        cua.nomcapat = Rs!nomcapat
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        chv.Tabla = "SCAPAT"
        chv.Oper = Tipo
        chv.fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = CStr(cua.codcapat)
        chv.XML = cua.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            chv.Clv_Old = chv.Clv_New
        End If
        chv.Grabar
    End If
End Sub



Sub CargarUnaPartida(codparti As Long, Tipo As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim par As GRPTC_Partida
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select * from rpartida "
    Sql = Sql & " where codparti = " & CStr(codparti)
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set par = New GRPTC_Partida
        '-- vamos cargando los diferentes valores
        par.codparti = Rs!codparti
        par.nomparti = Rs!nomparti
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        chv.Tabla = "SPARTI"
        chv.Oper = Tipo
        chv.fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = CStr(par.codparti)
        chv.XML = par.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            chv.Clv_Old = chv.Clv_New
        End If
        chv.Grabar
    End If
End Sub


Sub CargarUnVehiculo(codtrans As String, Tipo As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim tra As GRPTC_Vehiculo
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select * from rtransporte "
    Sql = Sql & " where codtrans = '" & codtrans & "'"
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set tra = New GRPTC_Vehiculo
        '-- vamos cargando los diferentes valores
        tra.nomcamio = Rs!nomtrans
        tra.matricul = Rs!codtrans
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        chv.Tabla = "SCAMIO"
        chv.Oper = Tipo
        chv.fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = CStr(tra.matricul)
        chv.XML = tra.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            chv.Clv_Old = chv.Clv_New
        End If
        chv.Grabar
    End If
End Sub



Sub CargarUnProducto(codprodu As Long, Tipo As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim pro As GRPTC_Producto
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select * from productos "
    Sql = Sql & " where codprodu = " & CStr(codprodu)
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set pro = New GRPTC_Producto
        '-- vamos cargando los diferentes valores
        pro.codprodu = Rs!codprodu
        pro.nomprodu = Rs!nomprodu
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        chv.Tabla = "SPRODU"
        chv.Oper = Tipo
        chv.fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = CStr(pro.codprodu)
        chv.XML = pro.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            chv.Clv_Old = chv.Clv_New
        End If
        chv.Grabar
    End If
End Sub

Sub CargarUnaVariedad(codvarie As Long, Tipo As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim vari As GRPTC_Variedad
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select * from variedades "
    Sql = Sql & " where codvarie = " & CStr(codvarie)
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set vari = New GRPTC_Variedad
        '-- vamos cargando los diferentes valores
        vari.codvarie = Rs!codvarie
        vari.nomvarie = Rs!nomvarie
        vari.codprodu = Rs!codprodu
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        chv.Tabla = "SVARIE"
        chv.Oper = Tipo
        chv.fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = CStr(vari.codvarie)
        chv.XML = vari.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            chv.Clv_Old = Rs!codprodu & "&" & Rs!codvarie 'chv.Clv_New
        End If
        chv.Grabar
    End If
End Sub



Public Function CApos(Texto As String) As String
    Dim i As Integer
    i = InStr(1, Texto, "'")
    If i = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, i - 1) & "\'" & Mid(Texto, i + 1, Len(Texto) - i)
    End If
    '-- Ya que estamos transformamos las Ñ
    Texto = CApos
    i = InStr(1, Texto, "¥")
    If i = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, i - 1) & "Ñ" & Mid(Texto, i + 1, Len(Texto) - i)
    End If
    '-- Y otra más
    Texto = CApos
    i = InStr(1, Texto, "¾")
    If i = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, i - 1) & "Ñ" & Mid(Texto, i + 1, Len(Texto) - i)
    End If
    '-- Seguimos con transformaciones
    Texto = CApos
    i = InStr(1, Texto, "¦")
    If i = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, i - 1) & "ª" & Mid(Texto, i + 1, Len(Texto) - i)
    End If
End Function

Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"
                    DBLet = ""
                Case "N"
                    DBLet = 0
                Case "F"
                    DBLet = CDate("0:00:00")
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
        If Tipo = "" Or Tipo = "T" Then DBLet = CStr(DBLet)
    End If
End Function


Public Function Cooperativa() As Integer
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "select cooperativa from rparam"
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        Cooperativa = DBLet(Rs.Fields(0).Value, "N")
    End If
    Set Rs = Nothing
    
End Function
