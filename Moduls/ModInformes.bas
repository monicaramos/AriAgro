Attribute VB_Name = "ModInformes"
Option Explicit


'==============================================================
'====== FUNCIONES GENERALES  PARA INFORMES ====================

'Esta funcion lo que hace es genera el valor del campo
'El campo lo coge del recordset, luego sera field(i), y el tipo es para a�adirle
'las coimllas, o quitarlas comas
'  Si es numero viene un 1 si no nada
'## NO LA USO, UTILIZO DBSET
'Public Function ParaBD(ByRef campo As ADODB.Field, Optional EsNumerico As Byte) As String
'
'    If IsNull(campo) Then
'        ParaBD = "NULL"
'    Else
'        Select Case EsNumerico
'        Case 1
'            ParaBD = TransformaComasPuntos(CStr(campo))
'        Case 2
'            'Fechas
'            ParaBD = "'" & Format(CStr(campo), "dd/MM/yyyy") & "'"
'        Case Else
'            ParaBD = "'" & campo & "'"
'        End Select
'    End If
'    ParaBD = "," & ParaBD
'End Function


Public Sub AbrirListado(numero As Byte)
    Screen.MousePointer = vbHourglass
    frmListado.Opcionlistado = numero
    frmListado.Show vbModal
    Screen.MousePointer = vbDefault
End Sub


Public Sub AbrirListadoOfer(numero As Integer)
'Abre el Form con los listados de Ofertas
    Screen.MousePointer = vbHourglass
    frmListadoOfer.Opcionlistado = numero
    frmListadoOfer.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Public Function AnyadirAFormula(ByRef cadFormula As String, arg As String) As Boolean
'Concatena los criterios del WHERE para pasarlos al Crystal como FormulaSelection
    If arg = "Error" Then
        AnyadirAFormula = False
        Exit Function
    ElseIf arg <> "" Then
        If cadFormula <> "" Then
            cadFormula = cadFormula & " AND " & arg
        Else
            cadFormula = arg
        End If
    End If
    AnyadirAFormula = True
End Function


Public Function RegistrosAListar(vSQL As String, Optional vBD As Byte) As Byte
'Devuelve si hay algun registro para mostrar en el Informe con la seleccion
'realizada. Si no hay nada que mostrar devuelve 0 y no abrir� el informe
Dim Rs As ADODB.Recordset

    On Error Resume Next
    
    Set Rs = New ADODB.Recordset
    If vBD = cConta Then
        Rs.Open vSQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Else
        Rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    End If

    
    RegistrosAListar = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then RegistrosAListar = 1 'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        RegistrosAListar = 0
        Err.Clear
    End If
End Function




Public Function HayRegParaInforme(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    If RegistrosAListar(SQL) = 0 Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegParaInforme = False
    Else
        HayRegParaInforme = True
    End If
End Function



Public Function CadenaDesdeHasta(cadDesde As String, cadHasta As String, campo As String, TipoCampo As String, Optional nomCampo As String) As String
'Devuelve la cadena de seleccion: " (campo >= cadDesde and campo<=cadHasta) "
'para Crystal Report
Dim cadAux As String

    If Trim(cadDesde) = "" And Trim(cadHasta) = "" Then
        'Campo Desde y Hasta no tiene valor
            cadAux = ""
    Else
        'Campo DESDE
        If cadDesde <> "" Then
            Select Case TipoCampo
                Case "N"
                    cadAux = campo & " >= " & Val(cadDesde)
                Case "T"
                    cadAux = campo & " >= """ & cadDesde & """"
                Case "F"
                    cadAux = campo & " >= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")"
            End Select
        End If
        
        'Campo HASTA
        If cadHasta <> "" Then
            If cadAux <> "" Then 'Hay campo Desde y campo Hasta
                'Comprobar Desde <= Hasta
                Select Case TipoCampo
                    Case "N"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= " & Val(cadHasta)
                        End If
                        
                    Case "T"
                        If cadDesde > cadHasta Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= """ & cadHasta & """"
                        End If
                    
                    Case "F"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= Date(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ")"
                        End If
                End Select
            Else 'No hay campo Desde. Solo hay campo Hasta
                Select Case TipoCampo
                    Case "N"
                        cadAux = campo & " <= " & Val(cadHasta)
                    Case "T"
                        cadAux = campo & " <= """ & cadHasta & """"
                    Case "F"
                        cadAux = campo & " <= Date(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ")"
                End Select
            End If
        End If
    End If
    If cadAux <> "" And cadAux <> "Error" Then cadAux = "(" & cadAux & ")"
    CadenaDesdeHasta = cadAux
End Function


Public Function CadenaDesdeHastaBD(cadDesde As String, cadHasta As String, campo As String, TipoCampo As String) As String
'Devuelve la cadena de seleccion: " (campo >= valor1 and campo<=valor2) "
'Para MySQL
Dim cadAux As String

    If Trim(cadDesde) = "" And Trim(cadHasta) = "" Then
        'Campo Desde y Hasta no tiene valor
            cadAux = ""
    Else
        'Campo DESDE
        If cadDesde <> "" Then
            Select Case TipoCampo
                Case "N"
                    cadAux = campo & " >= " & Val(cadDesde)
                Case "T"
                    cadAux = campo & " >= """ & cadDesde & """"
                Case "F"
                    cadAux = "(" & campo & " >= '" & Format(cadDesde, FormatoFecha) & "')"
            End Select
        End If
        
        'Campo HASTA
        If cadHasta <> "" Then
            If cadAux <> "" Then 'Hay campo Desde y campo Hasta
                'Comprobar Desde <= Hasta
                Select Case TipoCampo
                    Case "N"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= " & Val(cadHasta)
                        End If
                        
                    Case "T"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= """ & cadHasta & """"
                        End If
                    
                    Case "F"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and (" & campo & " <= '" & Format(cadHasta, FormatoFecha) & "')"
                        End If
                End Select
                
            Else 'No hay campo Desde. Solo hay campo Hasta
                Select Case TipoCampo
                    Case "N"
                        cadAux = campo & " <= " & Val(cadHasta)
                    Case "T"
                        cadAux = campo & " <= """ & cadHasta & """"
                    Case "F"
                        cadAux = campo & " <= '" & Format(cadHasta, FormatoFecha) & "'"
                End Select
            End If
        End If
    End If
    If cadAux <> "" And cadAux <> "Error" Then cadAux = "(" & cadAux & ")"
    CadenaDesdeHastaBD = cadAux
End Function


Public Function AnyadirParametroDH(param As String, codD As String, codH As String, nomD As String, nomH As String) As String
On Error Resume Next
    
    If codD <> "" Then
        param = param & "DESDE: " & codD
        If nomD <> "" Then param = param & " - " & Replace(nomD, """", """""") 'nomD
    End If
    If codH <> "" Then
        param = param & "  HASTA: " & codH
        If nomH <> "" Then param = param & " - " & Replace(nomH, """", """""") 'nomH
    End If
    
    AnyadirParametroDH = param & """|"
    If Err.Number <> 0 Then Err.Clear
End Function



Public Function QuitarCaracterACadena(cadForm As String, Caracter As String) As String
'IN: [cadForm] es la cadena en la que se eliminara todos los caractes iguales a la vble [Caracter]
'OUT: cadena sin los caracteres
'EJEMPLO: "{scaalb.numalbar}", "{"  -->  "scaalb.numalbar}"
Dim i As Integer
Dim J As Integer
Dim Aux As String

    Aux = cadForm
    i = InStr(1, Aux, Caracter, vbTextCompare)
    While i > 0
        i = InStr(1, Aux, Caracter, vbTextCompare)
        If i > 0 Then
            J = Len(Caracter)
            Aux = Mid(Aux, 1, i - 1) & Mid(Aux, i + J, Len(Aux) - 1)
        End If
    Wend
    QuitarCaracterACadena = Aux
End Function


'## Utilizo la funcion REPLACE
'Public Function SustituirCadenas(CADENA As String, cad1 As String, cad2 As String) As String
''IN: Cadena es la cadena de texto en la que se va a sustituir la cad1 por la cad2
''OUT: cadena con la sustitucion
'
''EJEMPLO: cadena = "scaalb.codtipom='ALV' AND scaalb.numalbar=1"
''         cad1 = "scaalb"
''         cad2 = "slialb"
'
''         Resultado = "slialb.codtipom='ALV' AND slialb.numalbar=1"
'
'Dim i As Integer
'Dim j As Integer
'Dim Aux As String
'
'    Aux = CADENA
'    Do
'        i = InStr(1, Aux, cad1, vbTextCompare)
'        If i > 0 Then
'            j = Len(cad1)
'            Aux = Mid(Aux, 1, i - 1) & cad2 & Mid(Aux, i + j, Len(Aux) - 1)
'        End If
'    Loop Until i <= 0
'    SustituirCadenas = Aux
'End Function




Public Function PonerParamRPT(indice As Byte, cadParam As String, numParam As Byte, nomDocu As String, Optional EsAridoc As Boolean, Optional ImprimeDirecto As Integer) As Boolean
'EsAridoc = false usamos el nomdocum normal
'           true usamos el rpt para aridoc
'ImprimeDirecto = false usamos el crystal
'                 true usamos el print

Dim vParamRpt As CParamRpt 'Tipos de Documentos
Dim cad As String

    Set vParamRpt = New CParamRpt

    If vParamRpt.Leer(indice) = 1 Then
        cad = "No se han podido cargar los Par�metros de Tipos de Documentos." & vbCrLf
        MsgBox cad & "Debe configurar la aplicaci�n.", vbExclamation
        Set vParamRpt = Nothing
        PonerParamRPT = False
        Exit Function
    Else
        If cadParam = "" Then
            cad = "|"
        Else
            cad = ""
        End If
        cad = cad & "pCodigoISO=""" & vParamRpt.CodigoISO & """|"
        If vParamRpt.CodigoRevision = -1 Then
            cad = cad & "pCodigoRev=""" & "" & """|"
        Else
            cad = cad & "pCodigoRev=""" & Format(vParamRpt.CodigoRevision, "00") & """|"
        End If
        numParam = numParam + 2
        If vParamRpt.LineaPie1 <> "" Then
            cad = cad & "pLinea1=""" & vParamRpt.LineaPie1 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie2 <> "" Then
            cad = cad & "pLinea2=""" & vParamRpt.LineaPie2 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie3 <> "" Then
            cad = cad & "pLinea3=""" & vParamRpt.LineaPie3 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie4 <> "" Then
            cad = cad & "pLinea4=""" & vParamRpt.LineaPie4 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie5 <> "" Then
            cad = cad & "pLinea5=""" & vParamRpt.LineaPie5 & """|"
            numParam = numParam + 1
        End If
        cadParam = cadParam & cad
        If Not EsAridoc Then
            nomDocu = vParamRpt.Documento
        Else
            nomDocu = vParamRpt.AridocRpt
        End If
        
        ImprimeDirecto = vParamRpt.ImprimeDirecto
        
        PonerParamRPT = True
        Set vParamRpt = Nothing
    End If
End Function



Public Sub PonerFrameVisible(ByRef vFrame As Frame, visible As Boolean, H As Integer, W As Integer)
'Pone el Frame Visible y Ajustado al Formulario, y visualiza los controles
    
        vFrame.visible = visible
        If visible = True Then
            'Ajustar Tama�o del Frame para ajustar tama�o de Formulario al del Frame
            vFrame.Top = -90
            vFrame.Left = 0
            vFrame.Width = W
            vFrame.Height = H
        End If
End Sub


Public Function PonerParamEmpresa(cadParam As String, numParam As Byte) As Boolean
Dim DomiEmp As String
Dim WebEmp As String
Dim cad As String

        DomiEmp = vParam.DomicilioEmpresa & " - " & vParam.CPostal & " " & vParam.Poblacion
        If vParam.Provincia <> vParam.Poblacion Then DomiEmp = DomiEmp & " " & vParam.Provincia
        DomiEmp = DomiEmp & " - Telf. " & vParam.Telefono & " - Fax. " & vParam.Fax
        WebEmp = "Internet: " & vParam.WebEmpresa & " - E-mail: " & vParam.MailEmpresa
        'Resto parametros
        cad = ""
        cad = cad & "pNomEmpre=""" & vParam.NombreEmpresa & """|"
        cad = cad & "pDomEmpre=""" & DomiEmp & """|"
        cad = cad & "pWebEmpre=""" & WebEmp & """|"
        
        numParam = numParam + 3
        cadParam = cadParam & cad
        PonerParamEmpresa = True
End Function

Public Function SaltosDeLinea(ByVal CADENA As String) As String
    Dim Devu As String
    Dim i As Integer
    
    Devu = ""
    Do
        i = InStr(1, CADENA, vbCrLf)
        If i > 0 Then
            If Devu <> "" Then Devu = Devu & """ + chr(13) + """
            Devu = Devu & Mid(CADENA, 1, i - 1)
            CADENA = Mid(CADENA, i + 2)
            
       Else
            Devu = Devu & CADENA
       End If
    Loop While i > 0
    SaltosDeLinea = Devu
End Function

