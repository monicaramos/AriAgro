Attribute VB_Name = "LibExport"
Option Explicit

Sub CargarTodosLosCampos()
    '-- Utilidad que carga todos los campos de la base de datos
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim cmp As GRPTC_Campo
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta �nica todos los campos
    SQL = "select a.*, b.codprodu from rcampos as a , variedades as b where b.codvarie = a.codvarie"
    Set RS = dbAriagro.cursor(SQL)
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            '-- creamos el objeto auxiliar que montar� el XML de trazatec
            Set cmp = New GRPTC_Campo
            '-- vamos cargando los diferentes valores
            cmp.CodSocio = RS!CodSocio
            cmp.CodCampo = RS!CodCampo
            cmp.codprodu = RS!codprodu
            cmp.codvarie = RS!codvarie
            cmp.codparti = RS!codparti
            cmp.hanegada = 0 ' no interesa en trazatec
            cmp.numarbol = 0 ' tampoco interesa
            cmp.poligono = RS!poligono
            '-- Y ahora el objeto chivato para grabar
            Set chv = New GRPTC_Chivato
            chv.Id = 0 'ya lo montar� en el momento de la grabaci�n
            chv.BD_Org = "AGRO"
            '[Monica]16/11/2011: Solo si es Alzira es SCAMP1
            If vParamAplic.Cooperativa = 4 Then
                chv.Tabla = "SCAMP1"
            Else
                chv.Tabla = "SCAMPO"
            End If
            chv.Oper = "I"
            chv.fecha = Format(Date, "dd/mm/yyyy")
            chv.Sep = "&"
            chv.Clv_Old = ""
            chv.Clv_New = CStr(cmp.CodCampo)
            chv.XML = cmp.GenXML
            chv.Grabar
            RS.MoveNext
        Wend
    End If
End Sub

Sub CargarUnCampo(CodCampo As Long, Tipo As String, Optional OldCadena As String)
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim cmp As GRPTC_Campo
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta �nica todos los campos
    SQL = "select a.*, b.codprodu from rcampos as a , variedades as b where b.codvarie = a.codvarie"
    SQL = SQL & " and a.codcampo = " & CStr(CodCampo)
    Set RS = dbAriagro.cursor(SQL)
    If Not RS.EOF Then
        '-- creamos el objeto auxiliar que montar� el XML de trazatec
        Set cmp = New GRPTC_Campo
        '-- vamos cargando los diferentes valores
        cmp.CodSocio = RS!CodSocio
        cmp.CodCampo = RS!CodCampo
        cmp.codprodu = RS!codprodu
        cmp.codvarie = RS!codvarie
        cmp.codparti = RS!codparti
        cmp.hanegada = 0 ' no interesa en trazatec
        cmp.numarbol = 0 ' tampoco interesa
        cmp.poligono = RS!poligono
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montar� en el momento de la grabaci�n
        chv.BD_Org = "AGRO"
        '[Monica]16/11/2011: solo en Alzira es la tabla SCAMP1
        If vParamAplic.Cooperativa = 4 Then
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
        
        chv.Clv_New = CStr(cmp.CodCampo)
        
        
        chv.XML = cmp.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            '[Monica]17/09/2013: solo para picassent cuando se est� modificando el campo
            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 7 Then
                chv.Clv_Old = OldCadena
            Else
                chv.Clv_Old = chv.Clv_New
            End If
        End If
        chv.Grabar
    End If
End Sub


Sub CargarUnSocio(CodSocio As Long, Tipo As String)
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim soc As GRPTC_Socio
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta �nica todos los campos
    SQL = "select * from rsocios "
    SQL = SQL & " where codsocio = " & CStr(CodSocio)
    Set RS = dbAriagro.cursor(SQL)
    If Not RS.EOF Then
        '-- creamos el objeto auxiliar que montar� el XML de trazatec
        Set soc = New GRPTC_Socio
        '-- vamos cargando los diferentes valores
        soc.CodSocio = RS!CodSocio
        soc.nifSocio = RS!nifSocio
        soc.NomSocio = RS!NomSocio
        soc.domsocio = RS!dirsocio
        soc.telsocio = DBLet(RS!telsoci1)
        soc.codPobla = 0 ' algo hay que hacer
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montar� en el momento de la grabaci�n
        chv.BD_Org = "AGRO"
        chv.Tabla = "SSOCIO"
        chv.Oper = Tipo
        chv.fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = CStr(soc.CodSocio)
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


Sub CargarUnaPoblacion(codPobla As String, Tipo As String)
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim pob As GRPTC_Poblacion
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta �nica todos los campos
    SQL = "select * from rpueblos "
    SQL = SQL & " where codpobla = '" & codPobla & "'"
    Set RS = dbAriagro.cursor(SQL)
    If Not RS.EOF Then
        '-- creamos el objeto auxiliar que montar� el XML de trazatec
        Set pob = New GRPTC_Poblacion
        '-- vamos cargando los diferentes valores
        pob.codPobla = RS!codPobla
        pob.desPobla = RS!desPobla
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montar� en el momento de la grabaci�n
        chv.BD_Org = "AGRO"
        chv.Tabla = "SPOBLA"
        chv.Oper = Tipo
        chv.fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = pob.codPobla
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
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim cua As GRPTC_Cuadrilla
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta �nica todos los campos
    SQL = "select * from rcapataz "
    SQL = SQL & " where codcapat = " & CStr(codcapat)
    Set RS = dbAriagro.cursor(SQL)
    If Not RS.EOF Then
        '-- creamos el objeto auxiliar que montar� el XML de trazatec
        Set cua = New GRPTC_Cuadrilla
        '-- vamos cargando los diferentes valores
        cua.codcapat = RS!codcapat
        cua.nomcapat = RS!nomcapat
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montar� en el momento de la grabaci�n
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
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim par As GRPTC_Partida
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta �nica todos los campos
    SQL = "select * from rpartida "
    SQL = SQL & " where codparti = " & CStr(codparti)
    Set RS = dbAriagro.cursor(SQL)
    If Not RS.EOF Then
        '-- creamos el objeto auxiliar que montar� el XML de trazatec
        Set par = New GRPTC_Partida
        '-- vamos cargando los diferentes valores
        par.codparti = RS!codparti
        par.nomparti = RS!nomparti
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montar� en el momento de la grabaci�n
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


Sub CargarUnVehiculo(codTrans As String, Tipo As String)
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim tra As GRPTC_Vehiculo
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta �nica todos los campos
    SQL = "select * from rtransporte "
    SQL = SQL & " where codtrans = '" & codTrans & "'"
    Set RS = dbAriagro.cursor(SQL)
    If Not RS.EOF Then
        '-- creamos el objeto auxiliar que montar� el XML de trazatec
        Set tra = New GRPTC_Vehiculo
        '-- vamos cargando los diferentes valores
        tra.nomcamio = RS!NomTrans
        tra.matricul = RS!codTrans
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montar� en el momento de la grabaci�n
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
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim pro As GRPTC_Producto
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta �nica todos los campos
    SQL = "select * from productos "
    SQL = SQL & " where codprodu = " & CStr(codprodu)
    Set RS = dbAriagro.cursor(SQL)
    If Not RS.EOF Then
        '-- creamos el objeto auxiliar que montar� el XML de trazatec
        Set pro = New GRPTC_Producto
        '-- vamos cargando los diferentes valores
        pro.codprodu = RS!codprodu
        pro.nomprodu = RS!nomprodu
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montar� en el momento de la grabaci�n
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

Sub CargarUnaVariedad(codvarie As Long, Tipo As String, Optional OldCadena As String)
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim vari As GRPTC_Variedad
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta �nica todos los campos
    SQL = "select * from variedades "
    SQL = SQL & " where codvarie = " & CStr(codvarie)
    Set RS = dbAriagro.cursor(SQL)
    If Not RS.EOF Then
        '-- creamos el objeto auxiliar que montar� el XML de trazatec
        Set vari = New GRPTC_Variedad
        '-- vamos cargando los diferentes valores
        vari.codvarie = RS!codvarie
        vari.nomvarie = RS!nomvarie
        vari.codprodu = RS!codprodu
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montar� en el momento de la grabaci�n
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
            '[Monica]18/09/2013: Si es Picassent o Quatretonda tengo que meter producto variedad
            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 7 Then
                chv.Clv_Old = OldCadena
            Else
                chv.Clv_Old = chv.Clv_New
            End If
        End If
        chv.Grabar
    End If
End Sub



