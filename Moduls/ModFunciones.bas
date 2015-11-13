Attribute VB_Name = "ModFunciones"
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
'   En este modulo estan las funciones que recorren el form
'   usando el each for
'   Estas son
'
'   CamposSiguiente -> Nos devuelve el el text siguiente en
'           el orden del tabindex
'
'   CompForm -> Compara los valores con su tag
'
'   InsertarDesdeForm - > Crea el sql de insert e inserta
'
'   Limpiar -> Pone a "" todos los objetos text de un form
'
'   ObtenerBusqueda -> A partir de los text crea el sql a
'       partir del WHERE ( sin el).
'
'   ModifcarDesdeFormulario -> Opcion modificar. Genera el SQL
'
'   PonerDatosForma -> Pone los datos del RECORDSET en el form
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
Option Explicit


Public Const ValorNulo = "Null"
Public NombreCheck As String


Public Function CompForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    CompForm = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox And Control.visible = True Then
            Carga = mTag.Cargar(Control)
            If Carga = True Then
                Correcto = mTag.Comprobar(Control)
                If Not Correcto Then Exit Function
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga = False Then
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function

                Else
                    If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                            MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                            Exit Function
                    End If
                End If
            End If
        End If
    Next Control
    CompForm = True
End Function

'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function CompForm2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    CompForm2 = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox Then
            If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                Carga = mTag.Cargar(Control)
                If Carga = True Then
                    Correcto = mTag.Comprobar(Control)
                    If Not Correcto Then Exit Function
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    Carga = mTag.Cargar(Control)
                    If Carga = False Then
                        MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                        Exit Function
    
                    Else
                        If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                                MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                                Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next Control
    CompForm2 = True
End Function




'Public Function CampoSiguiente(ByRef formulario As Form, valor As Integer) As Control
'Dim Fin As Boolean
'Dim Control As Object
'
'On Error GoTo ECampoSiguiente
'
'    'Debug.Print "Llamada:  " & Valor
'    'Vemos cual es el siguiente
'    Do
'        valor = valor + 1
'        For Each Control In formulario.Controls
'            'Debug.Print "-> " & Control.Name & " - " & Control.TabIndex
'            'Si es texto monta esta parte de sql
'            If Control.TabIndex = valor Then
'                    Set CampoSiguiente = Control
'                    Fin = True
'                    Exit For
'            End If
'        Next Control
'        If Not Fin Then
'            valor = -1
'        End If
'    Loop Until Fin
'    Exit Function
'ECampoSiguiente:
'    Set CampoSiguiente = Nothing
'    Err.Clear
'End Function



'-----------------------------------
Public Function ValorParaSQL(Valor, ByRef vtag As CTag) As String
Dim Dev As String
Dim d As Single
Dim I As Integer
Dim V
    Dev = ""
    If Valor <> "" Then
        Select Case vtag.TipoDato
        Case "N"
            V = Valor
            If InStr(1, Valor, ",") Or InStr(1, Valor, ".") Then
                If InStr(1, Valor, ".") Then
                    'ABRIL 2004

                    'Ademas de la coma lleva puntos
                    V = ImporteFormateado(CStr(Valor))
                    Valor = V
                Else

                    V = CSng(Valor)
                    Valor = V
                End If
            Else

            End If
            Dev = TransformaComasPuntos(CStr(Valor))

        Case "F"
            Dev = "'" & Format(Valor, FormatoFecha) & "'"
            
        Case "H"
            Dev = "'" & Format(Valor, "hh:mm:ss") & "'"
        
        Case "FHH"
            Dev = DBSet(Valor, "FH")
            
        Case "FH"
            Dev = DBSet(Valor, "FH")
        Case "T"
            Dev = CStr(Valor)
            NombreSQL Dev
            Dev = "'" & Dev & "'"
        Case Else
            Dev = "'" & Valor & "'"
        End Select

    Else
        'Si se permiten nulos, la "" ponemos un NULL
        If vtag.Vacio = "S" Then Dev = ValorNulo
    End If
    ValorParaSQL = Dev
End Function


Public Function InsertarDesdeForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim cad As String
    
    On Error GoTo EInsertarF
    
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm = False
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                    
                        'Parte VALUES
                        cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.columna & ""
                If Control.Value = 1 Then
                    cad = "1"
                    Else
                    cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                Der = Der & cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.ListIndex = -1 Then
                        cad = ValorNulo
                    Else
                        cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "INSERT INTO " & mTag.tabla
    cad = cad & " (" & Izda & ") VALUES (" & Der & ");"
    
    conn.Execute cad, , adCmdText
    
    InsertarDesdeForm = True
Exit Function

EInsertarF:
    MuestraError Err.Number, "Inserta. " & Err.Description
End Function


Public Function InsertarDesdeForm2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim cad As String
    
    On Error GoTo EInsertarF
    
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm2 = False
    Der = ""
    Izda = ""
    
    For Each Control In formulario.Controls
    
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            If Izda <> "" Then Izda = Izda & ","
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.columna & ""
                        
                            'Parte VALUES
                            cad = ValorParaSQL(Control.Text, mTag)
                            If Der <> "" Then Der = Der & ","
                            Der = Der & cad
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Izda <> "" Then Izda = Izda & ","
                    'Access
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.Value = 1 Then
                        cad = "1"
                        Else
                        cad = "0"
                    End If
                    If Der <> "" Then Der = Der & ","
                    If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                    Der = Der & cad
                End If
            End If
            
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                        If Control.ListIndex = -1 Then
                            cad = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            cad = Control.ItemData(Control.ListIndex)
                        Else
                            cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
            
        'OPTION BUTTON
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            If Izda <> "" Then Izda = Izda & ","
                            Izda = Izda & "" & mTag.columna & ""
                            cad = Control.Index
                            If Der <> "" Then Der = Der & ","
                            Der = Der & cad
                        End If
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is DTPicker Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
'                        If Control.Value Then
'                            If Izda <> "" Then Izda = Izda & ","
'                            Izda = Izda & "" & mTag.columna & ""
'                            cad = Control.index
'                            If Der <> "" Then Der = Der & ","
'                            Der = Der & cad
'                        End If
                        If Izda <> "" Then Izda = Izda & ","
                        Izda = Izda & "" & mTag.columna & ""
                        
                        'Parte VALUES
                        If Control.visible Then
                            cad = ValorParaSQL(Control.Value, mTag)
                        Else
                            cad = ValorNulo
                        End If
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
        End If
        
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "INSERT INTO " & mTag.tabla
    cad = cad & " (" & Izda & ") VALUES (" & Der & ");"
    conn.Execute cad, , adCmdText
    
     ' ### [Monica] 18/12/2006
    CadenaCambio = cad
   
    InsertarDesdeForm2 = True
Exit Function

EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function


Public Function CadenaInsertarDesdeForm(ByRef formulario As Form) As String
'Equivale a InsertarDesdeForm, excepto que devuelve la candena SQL y hace el execute fuera de la función.
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    CadenaInsertarDesdeForm = ""
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                    
                        'Parte VALUES
                        cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.columna & ""
                If Control.Value = 1 Then
                    cad = "1"
                    Else
                    cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                Der = Der & cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.ListIndex = -1 Then
                        cad = ValorNulo
                    Else
                        cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "INSERT INTO " & mTag.tabla
    cad = cad & " (" & Izda & ") VALUES (" & Der & ");"
'    Conn.Execute cad, , adCmdText
    
    CadenaInsertarDesdeForm = cad
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function



Public Function PonerCamposForma(ByRef formulario As Form, ByRef vData As Adodc) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim cad As String
Dim Valor As Variant
Dim campo As String  'Campo en la base de datos
Dim I As Integer

    Set mTag = New CTag
    PonerCamposForma = False

    For Each Control In formulario.Controls
        'TEXTO
        If (TypeOf Control Is TextBox) And (Control.visible = True) And UCase(Control.Name) = "TEXT1" Then
'                If TypeOf control Is TextBox Then

            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    If mTag.columna <> "" Then
                        campo = mTag.columna
                        If mTag.Vacio = "S" Then
                            Valor = DBLet(vData.Recordset.Fields(campo))
                        Else
                            Valor = vData.Recordset.Fields(campo)
                        End If
                        If mTag.Formato <> "" And CStr(Valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                cad = Format(Valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = cad
                            Else
                                Control.Text = Format(Valor, mTag.Formato)
                            End If
                        Else
                            Control.Text = Valor
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf (TypeOf Control Is CheckBox) And (Control.visible = True) Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    campo = mTag.columna
                    Valor = vData.Recordset.Fields(campo)
                    Else
                        Valor = 0
                End If
                If IsNull(Valor) Then Valor = 0
                Control.Value = Valor
            End If

         'COMBOBOX
         ElseIf (TypeOf Control Is ComboBox) And Control.visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    campo = mTag.columna
                    Valor = DBLet(vData.Recordset.Fields(campo))
                    '++MONICA: 15-01-2008 añadida la condicion de que el valor sea nulo
                    If IsNull(vData.Recordset.Fields(campo)) Then
                        Control.ListIndex = -1
                    Else
                    '++
                        I = 0
                        For I = 0 To Control.ListCount - 1
                            If Control.ItemData(I) = Val(Valor) Then
                                Control.ListIndex = I
                                Exit For
                            End If
                        Next I
                        If I = Control.ListCount Then Control.ListIndex = -1
                    '++MONICA: 15-01-2008 añadida la condicion de que el valor sea nulo
                    End If
                    '++
                End If 'de cargado
            End If 'de <>""
        End If
    Next Control

    'Veremos que tal
    PonerCamposForma = True
Exit Function
EPonerCamposForma:
    MuestraError Err.Number, "Poner campos formulario. "
End Function

'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function PonerCamposForma2(ByRef formulario As Form, ByRef vData As Adodc, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim cad As String
Dim Valor As Variant
Dim campo As String  'Campo en la base de datos
Dim I As Integer
    Set mTag = New CTag
    PonerCamposForma2 = False

    For Each Control In formulario.Controls
        'TEXTO
        If (TypeOf Control Is TextBox) Then
            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    If mTag.Cargado Then
                        'Columna en la BD
                        If mTag.columna <> "" Then
                            campo = mTag.columna
                            If mTag.Vacio = "S" Then
                                Valor = DBLet(vData.Recordset.Fields(campo))
                            Else
                                Valor = vData.Recordset.Fields(campo)
                            End If
                            If mTag.Formato <> "" And CStr(Valor) <> "" Then
                                If mTag.TipoDato = "N" Then
                                    'Es numerico, entonces formatearemos y sustituiremos
                                    ' La coma por el punto
                                    cad = Format(Valor, mTag.Formato)
                                    'Antiguo
                                    'Control.Text = TransformaComasPuntos(cad)
                                    'nuevo
                                    Control.Text = cad
                                Else
                                    Control.Text = Format(Valor, mTag.Formato)
                                End If
                            Else
                                Control.Text = Valor
                            End If
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf (TypeOf Control Is CheckBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Columna en la BD
                        campo = mTag.columna
                        Valor = vData.Recordset.Fields(campo)
                    Else
                        Valor = 0
                    End If
                    If IsNull(Valor) Then Valor = 0
                    Control.Value = Valor
                End If
            End If

         'COMBOBOX
         ElseIf (TypeOf Control Is ComboBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        campo = mTag.columna
                        Valor = DBLet(vData.Recordset.Fields(campo))
                        I = 0
                        For I = 0 To Control.ListCount - 1
                            If Control.ItemData(I) = Val(Valor) Then
                                Control.ListIndex = I
                                Exit For
                            End If
                        Next I
                        If I = Control.ListCount Then Control.ListIndex = -1
                    End If 'de cargado
                End If
            End If 'de <>""
            
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Columna en la BD
                        campo = mTag.columna
                        Valor = vData.Recordset.Fields(campo)
                        If IsNull(Valor) Then Valor = 0
                        If Control.Index = Valor Then
                            Control.Value = True
                        Else
                            Control.Value = False
                        End If
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is DTPicker Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Columna en la BD
                        campo = mTag.columna
                        Valor = vData.Recordset.Fields(campo)
                        If IsNull(Valor) Then Valor = Now
                        Control.Value = Format(Valor, mTag.Formato)
                    End If
                End If
            End If
        End If
    Next Control

    'Veremos que tal
    PonerCamposForma2 = True
Exit Function
EPonerCamposForma2:
    MuestraError Err.Number, "Poner campos formulario 2. "
End Function


Public Function ForaGrid(ByRef formulari As Form, ByRef vGrid As DataGrid, Control As Object) As Boolean
Dim mTag As CTag
Dim cad As String
Dim Valor As Variant
Dim camp As String  'Camp en la BDA
Dim I As Integer

    Set mTag = New CTag
    ForaGrid = False

    If (TypeOf Control Is TextBox) Then 'text
        mTag.Cargar Control
        If Control.Tag <> "" Then
            If mTag.Cargado Then
                If mTag.columna <> "" Then
                    camp = mTag.columna
                    If mTag.Vacio = "S" Then
                        Valor = DBLet(vGrid.Columns(camp).Text)
                        'valor = DBLet(vGrid.Recordset.Fields(campo))
                    Else
                        'valor = vGrid.Columns!camp
                        Valor = vGrid.Columns(camp).Text
                    End If
                    If mTag.Formato <> "" And CStr(Valor) <> "" Then
                        If mTag.TipoDato = "N" Then
                            cad = Format(Valor, mTag.Formato)
                            Control.Text = cad
                        Else
                            Control.Text = Format(Valor, mTag.Formato)
                        End If
                    Else
                        Control.Text = Valor
                    End If
                End If
            End If
        End If

'        'CheckBOX
'        ElseIf (TypeOf Control Is CheckBox) Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado Then
'                        'Columna en la BD
'                        campo = mTag.columna
'                        valor = vData.Recordset.Fields(campo)
'                        Else
'                            valor = 0
'                    End If
'                    If IsNull(valor) Then valor = 0
'                    Control.Value = valor
'                End If
'            End If
'
'         'COMBOBOX
'         ElseIf (TypeOf Control Is ComboBox) Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado Then
'                        campo = mTag.columna
'                        valor = DBLet(vData.Recordset.Fields(campo))
'                        i = 0
'                        For i = 0 To Control.ListCount - 1
'                            If Control.ItemData(i) = Val(valor) Then
'                                Control.ListIndex = i
'                                Exit For
'                            End If
'                        Next i
'                        If i = Control.ListCount Then Control.ListIndex = -1
'                    End If 'de cargado
'                End If
'            End If 'de <>""
    End If

    'Veremos que tal
    ForaGrid = True
Exit Function
EPosarCampsGrid:
    MuestraError Err.Number, "Poner campos grid. "
End Function


'Public Function PonerCamposFormaFrame(ByRef formulario As Form, NomTxtBox As String, ByRef vData As Adodc, Optional NomCheck As String, Optional NomCombo As String) As Boolean
'Dim Control As Object
'Dim mTag As CTag
'Dim cad As String
'Dim valor As Variant
'Dim campo As String  'Campo en la base de datos
'Dim i As Integer
'
'    Set mTag = New CTag
'    PonerCamposFormaFrame = False
'
'
'        For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox And Control.Visible = True And Control.Name = NomTxtBox Then
'            'Comprobamos que tenga tag
'            mTag.Cargar Control
''            Debug.Print Control.Parent
'            If Control.Tag <> "" Then
'                If mTag.Cargado Then
'                    'Columna en la BD
'                    If mTag.Columna <> "" Then
'                        campo = mTag.Columna
'                        If mTag.Vacio = "S" Then
'                            valor = DBLet(vData.Recordset.Fields(campo))
'                        Else
'                            valor = vData.Recordset.Fields(campo)
'                        End If
'                        If mTag.Formato <> "" And CStr(valor) <> "" Then
'                            If mTag.TipoDato = "N" Then
'                                'Es numerico, entonces formatearemos y sustituiremos
'                                ' La coma por el punto
'                                cad = Format(valor, mTag.Formato)
'                                'Antiguo
'                                'Control.Text = TransformaComasPuntos(cad)
'                                'nuevo
'                                Control.Text = cad
'                            Else
'                                Control.Text = Format(valor, mTag.Formato)
'                            End If
'                        Else
'                            Control.Text = valor
'                        End If
'                    End If
'                End If
'            End If
'        'CheckBOX
'        ElseIf TypeOf Control Is CheckBox And Control.Visible = True And Control.Name = NomCheck Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    'Columna en la BD
'                    campo = mTag.Columna
'                    valor = vData.Recordset.Fields(campo)
'                    Else
'                        valor = 0
'                End If
'                Control.Value = valor
'            End If
'
'         'COMBOBOX
'         ElseIf TypeOf Control Is ComboBox And Control.Visible = True And Control.Name = NomCombo Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    campo = mTag.Columna
'                    valor = vData.Recordset.Fields(campo)
'                    i = 0
'                    For i = 0 To Control.ListCount - 1
'                        If Control.ItemData(i) = Val(valor) Then
'                            Control.ListIndex = i
'                            Exit For
'                        End If
'                    Next i
'                    If i = Control.ListCount Then Control.ListIndex = -1
'                End If 'de cargado
'            End If 'de <>""
'        End If
'
'    Next Control
'
'    'Veremos que tal
'    PonerCamposFormaFrame = True
'Exit Function
'EPonerCamposForma:
'    MuestraError Err.Number, "Poner campos formulario. "
'End Function


Private Function ObtenerMaximoMinimo(vSQL As String, Optional vBD As Byte) As String
Dim Rs As Recordset
    ObtenerMaximoMinimo = ""
    Set Rs = New ADODB.Recordset
    If vBD = cConta Then
        Rs.Open vSQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else
        Rs.Open vSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            ObtenerMaximoMinimo = CStr(Rs.Fields(0))
        End If
    End If
    Rs.Close
    Set Rs = Nothing
End Function


'====DAVID
'Public Function ObtenerBusqueda(ByRef formulario As Form) As String
'    Dim Control As Object
'    Dim Carga As Boolean
'    Dim mTag As CTag
'    Dim Aux As String
'    Dim cad As String
'    Dim SQL As String
'    Dim tabla As String
'    Dim RC As Byte
'
'    On Error GoTo EObtenerBusqueda
'
'    'Exit Function
'    Set mTag = New CTag
'    ObtenerBusqueda = ""
'    SQL = ""
'
'    'Recorremos los text en busca de ">>" o "<<"
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If Aux = ">>" Or Aux = "<<" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'                    If Aux = ">>" Then
'                        cad = " MAX(" & mTag.Columna & ")"
'                    Else
'                        cad = " MIN(" & mTag.Columna & ")"
'                    End If
'                    SQL = "Select " & cad & " from " & mTag.tabla
'                    SQL = ObtenerMaximoMinimo(SQL)
'                    Select Case mTag.TipoDato
'                    Case "N"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(SQL)
'                    Case "F"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
'                    Case Else
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & SQL & "'"
'                    End Select
'                    SQL = "(" & SQL & ")"
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los text en busca del NULL
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If UCase(Aux) = "NULL" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'
'                    SQL = mTag.tabla & "." & mTag.Columna & " is NULL"
'                    SQL = "(" & SQL & ")"
'                    Control.Text = ""
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los textbox
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            'Cargamos el tag
'            Carga = mTag.Cargar(Control)
'            If Carga Then
'                If mTag.Cargado Then
'                    Aux = Trim(Control.Text)
'                    If Aux <> "" Then
'                        If mTag.tabla <> "" Then
'                            tabla = mTag.tabla & "."
'                        Else
'                            tabla = ""
'                        End If
'                    RC = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.Columna, Aux, cad)
'                    If RC = 0 Then
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'            Else
'                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
'                Exit Function
'            End If
'
'        'COMBO BOX
'        ElseIf TypeOf Control Is ComboBox Then
'            mTag.Cargar Control
'            If mTag.Cargado Then
'                If Control.ListIndex > -1 Then
'                    If mTag.TipoDato <> "T" Then
'                        cad = Control.ItemData(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = " & cad
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    Else
'                        cad = Control.List(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = '" & cad & "'"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'
'
'        'CHECK
'        ElseIf TypeOf Control Is CheckBox Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    If Control.Value = 1 Then
'                        cad = mTag.tabla & "." & mTag.Columna & " = 1"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'        End If
'
'
'    Next Control
'    ObtenerBusqueda = SQL
'Exit Function
'EObtenerBusqueda:
'    ObtenerBusqueda = ""
'    MuestraError Err.Number, "Obtener búsqueda. "
'End Function

Public Function ObtenerBusqueda(ByRef formulario As Form, Optional CHECK As String, Optional vBD As Byte, Optional cadWHERE As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim cad As String
    Dim Sql As String
    Dim tabla As String
    Dim Rc As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                If Control.Tag <> "" Then
                    Carga = mTag.Cargar(Control)
                    If Carga Then
                        If Aux = ">>" Then
                            cad = " MAX("
                        Else
                            cad = " MIN("
                        End If
                        'monica
                        Select Case mTag.TipoDato
                            Case "FHF"
                                cad = cad & "date(" & mTag.columna & "))"
                            Case "FHH"
                                cad = cad & "time(" & mTag.columna & "))"
                            Case Else
                                cad = cad & mTag.columna & ")"
                        End Select
                        
                        Sql = "Select " & cad & " from " & mTag.tabla
                        If cadWHERE <> "" Then Sql = Sql & " WHERE " & cadWHERE
                        Sql = ObtenerMaximoMinimo(Sql, vBD)
                        Select Case mTag.TipoDato
                        Case "N"
                            Sql = mTag.tabla & "." & mTag.columna & " = " & TransformaComasPuntos(Sql)
                        Case "F"
                            Sql = mTag.tabla & "." & mTag.columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case "FHF"
                            Sql = "date(" & mTag.tabla & "." & mTag.columna & ") = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case "FHH"
                            Sql = "time(" & mTag.tabla & "." & mTag.columna & ") = '" & Format(Sql, "hh:mm:ss") & "'"
                        Case Else
                            Sql = mTag.tabla & "." & mTag.columna & " = '" & Sql & "'"
                        End Select
                        Sql = "(" & Sql & ")"
                    End If
                End If
            End If
        End If
    Next


'++monica: lo he añadido del anterior obtenerbusqueda
    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then

                    Sql = mTag.tabla & "." & mTag.columna & " is NULL"
                    Sql = "(" & Sql & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next
 

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                'Cargamos el tag
                Carga = mTag.Cargar(Control)
                If Carga Then
'                    Debug.Print Control.Tag
                    Aux = Trim(Control.Text)
                    If Aux <> "" Then
                        If mTag.tabla <> "" Then
                            tabla = mTag.tabla & "."
                            Else
                            tabla = ""
                        End If
                        Rc = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.columna, Aux, cad)
                        If Rc = 0 Then
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex > -1 Then
                        If mTag.TipoDato = "N" Then
                            cad = Control.ItemData(Control.ListIndex)
                        Else
                            cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        cad = mTag.tabla & "." & mTag.columna & " = " & cad
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is CheckBox Then
            '=============== Añade: Laura, 15/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    Aux = ""
                    If CHECK <> "" Then
                        tabla = DBLet(Control.Index, "T")
                        If tabla <> "" Then tabla = "(" & tabla & ")"
                        tabla = Control.Name & tabla & "|"
                        If InStr(1, CHECK, tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    If Aux <> "" Then
'                    If Control.Value = 1 Then
                        cad = Control.Value
                        cad = mTag.tabla & "." & mTag.columna & " = " & cad
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusqueda = Sql
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener búsqueda. " & vbCrLf & Err.Description
End Function

'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function ObtenerBusqueda2(ByRef formulario As Form, Optional CHECK As String, Optional opcio As Integer, Optional nom_frame As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim cad As String
    Dim Sql As String
    Dim tabla As String
    Dim Rc As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda2 = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Aux = ">>" Then
                            cad = " MAX(" & mTag.columna & ")"
                        Else
                            cad = " MIN(" & mTag.columna & ")"
                        End If
                        Sql = "Select " & cad & " from " & mTag.tabla
                        Sql = ObtenerMaximoMinimo(Sql)
                        Select Case mTag.TipoDato
                        Case "N"
                            Sql = mTag.tabla & "." & mTag.columna & " = " & TransformaComasPuntos(Sql)
                        Case "F"
                            Sql = mTag.tabla & "." & mTag.columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case Else
                            Sql = mTag.tabla & "." & mTag.columna & " = '" & Sql & "'"
                        End Select
                        Sql = "(" & Sql & ")"
                    End If
                End If
            End If
        End If
    Next

'++monica: lo he añadido del anterior obtenerbusqueda
    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then

                    Sql = mTag.tabla & "." & mTag.columna & " is NULL"
                    Sql = "(" & Sql & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
          If Control.Tag <> "" Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    Aux = Trim(Control.Text)
                    If Aux <> "" Then
                        If mTag.tabla <> "" Then
                            tabla = mTag.tabla & "."
                            Else
                            tabla = ""
                        End If
                        Rc = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.columna, Aux, cad)
                        If Rc = 0 Then
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        End If
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then ' +-+- 12/05/05: canvi de Cèsar, no te sentit passar-li un control que no té TAG +-+-
                mTag.Cargar Control
                If mTag.Cargado Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Control.ListIndex > -1 Then
                            cad = Control.ItemData(Control.ListIndex)
                            cad = mTag.tabla & "." & mTag.columna & " = " & cad
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                End If
            End If
            
         ElseIf TypeOf Control Is CheckBox Then
            '=============== Añade: Laura, 27/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    ' añadido 12022007
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    
                        Aux = ""
                        If CHECK <> "" Then
                            tabla = DBLet(Control.Index, "T")
                            If tabla <> "" Then tabla = "(" & tabla & ")"
                            tabla = Control.Name & tabla & "|"
                            If InStr(1, CHECK, tabla, vbTextCompare) > 0 Then Aux = Control.Value
                        Else
                            If Control.Value = 1 Then Aux = "1"
                        End If
                        If Aux <> "" Then
    '                    If Control.Value = 1 Then
                            cad = Control.Value
                            cad = mTag.tabla & "." & mTag.columna & " = " & cad
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & cad & ")"
                        End If
                        
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusqueda2 = Sql
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda2 = ""
    MuestraError Err.Number, "Obtener búsqueda. " & vbCrLf & Err.Description
End Function

'Añado Optional CHECK As String. Para poder realizar las busquedas con los checks
'monica corresponde al ObtenerBusqueda de laura
Public Function ObtenerBusqueda3(ByRef formulario As Form, paraRPT As Boolean, Optional CHECK As String) As String
Dim Control As Object
Dim Carga As Boolean
Dim mTag As CTag
Dim Aux As String
Dim cad As String
Dim Sql As String
Dim tabla As String, columna As String
Dim Rc As Byte

    On Error GoTo EObtenerBusqueda3

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda3 = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.visible Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Aux = ">>" Then
                        If Not paraRPT Then
                            cad = " MAX(" & mTag.columna & ")"
                        Else
                            cad = " MAX({" & mTag.tabla & "." & mTag.columna & "})"
                        End If
                    Else
                        If Not paraRPT Then
                            cad = " MIN(" & mTag.columna & ")"
                        Else
                            cad = " MIN({" & mTag.tabla & "." & mTag.columna & "})"
                        End If
                    End If
                    If Not paraRPT Then
                        Sql = "Select " & cad & " from " & mTag.tabla
                    Else
                        Sql = "Select " & cad & " from {" & mTag.tabla & "}"
                    End If
                    Sql = ObtenerMaximoMinimo(Sql)
                    
                    Select Case mTag.TipoDato
                    Case "N"
                        If Not paraRPT Then
                            Sql = mTag.tabla & "." & mTag.columna & " = " & TransformaComasPuntos(Sql)
                        Else
                            Sql = "{" & mTag.tabla & "." & mTag.columna & "} = " & TransformaComasPuntos(Sql)
                        End If
                    Case "F"
                        If Not paraRPT Then
                            Sql = mTag.tabla & "." & mTag.columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Else
                            Sql = "{" & mTag.tabla & "." & mTag.columna & "} = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        End If
                    Case Else
                        If Not paraRPT Then
                            Sql = mTag.tabla & "." & mTag.columna & " = '" & Sql & "'"
                        Else
                            Sql = "{" & mTag.tabla & "." & mTag.columna & "} = '" & Sql & "'"
                        End If
                    End Select
                    Sql = "(" & Sql & ")"
                End If
            End If
        End If
    Next

    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.visible Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Not paraRPT Then
                        Sql = mTag.tabla & "." & mTag.columna & " is NULL"
                    Else
                        Sql = "{" & mTag.tabla & "." & mTag.columna & "} is NULL"
                    End If
                    Sql = "(" & Sql & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.visible Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                If mTag.Cargado Then
                    Aux = Trim(Control.Text)
                    Aux = QuitarCaracterEnter(Aux) 'Si es multilinea quitar ENTER
                    If Aux <> "" Then
                        If mTag.tabla <> "" Then
                            If Not paraRPT Then
                                tabla = mTag.tabla & "."
                            Else
                                tabla = "{" & mTag.tabla & "."
                            End If
                        Else
                            tabla = ""
                        End If
                        If Not paraRPT Then
                            columna = mTag.columna
                        Else
                            columna = mTag.columna & "}"
                        End If
                    Rc = SeparaCampoBusqueda3(mTag.TipoDato, tabla & columna, Aux, cad, paraRPT)
                    If Rc = 0 Then
                        If Sql <> "" Then Sql = Sql & " AND "
                        If Not paraRPT Then
                            Sql = Sql & "(" & cad & ")"
                        Else
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                End If
            End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If

        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.visible Then
            mTag.Cargar Control
            If mTag.Cargado Then
                If Control.ListIndex > -1 Then
                    If mTag.TipoDato <> "T" Then
                        cad = Control.ItemData(Control.ListIndex)
                        If Not paraRPT Then
                            cad = mTag.tabla & "." & mTag.columna & " = " & cad
                        Else
                            cad = "{" & mTag.tabla & "." & mTag.columna & "} = " & cad
                        End If
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    Else
                        cad = Control.List(Control.ListIndex)
                        If Not paraRPT Then
                            cad = mTag.tabla & "." & mTag.columna & " = '" & cad & "'"
                        Else
                            cad = "{" & mTag.tabla & "." & mTag.columna & "} = '" & cad & "'"
                        End If
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If


        'CHECK
                'CHECK
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    
                    Aux = ""
                    If CHECK <> "" Then
                        CheckBusqueda Control
                        tabla = NombreCheck & "|"
                        If InStr(1, CHECK, tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    
                    If Aux <> "" Then
                        If Not paraRPT Then
                            cad = mTag.tabla & "." & mTag.columna
                        Else
                            cad = "{" & mTag.tabla & "." & mTag.columna & "} "
                        End If
                        
                        cad = cad & " = " & Aux
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If 'cargado
                End If '<>""
            End If
        End If
    
    Next Control
    ObtenerBusqueda3 = Sql
Exit Function
EObtenerBusqueda3:
    ObtenerBusqueda3 = ""
    MuestraError Err.Number, "Obtener búsqueda. "
End Function

Public Function ModificaDesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        'Sea para el where o para el update esto lo necesito
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                             cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"

                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox And Control.visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        'Lo pondremos para el WHERE
                         If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                         cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                    Else
                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
'
'
'                   If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                   'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
'                   cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    conn.Execute Aux, , adCmdText

    ModificaDesdeFormulario = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function


Public Function ModificaDesdeFormulario2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario2 = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                 If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                                 cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
    
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Control.Value = 1 Then
                        Aux = "TRUE"
                    Else
                        Aux = "FALSE"
                    End If
                    If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'Esta es para access
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If

        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.ListIndex = -1 Then
                            Aux = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            Aux = Control.ItemData(Control.ListIndex)
                        Else
                            Aux = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                             cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                        End If
'
'
'                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                        'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
'                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            Aux = Control.Index
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                              If mTag.EsClave Then
                                  'Lo pondremos para el WHERE
                                   If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                                   cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                              Else
                                  If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                  cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                              End If
                        End If
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is DTPicker Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
'                        If Control.Value Then
                         If mTag.columna <> "" Then
'                            Aux = Control.index
                            If Control.visible Then
                                Aux = ValorParaSQL(Control.Value, mTag)
                            Else
                                Aux = ValorNulo
                            End If
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                                cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    conn.Execute Aux, , adCmdText

    ' ### [Monica] 18/12/2006
    CadenaCambio = cadUPDATE

    ModificaDesdeFormulario2 = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar 2. " & Err.Description
End Function

Public Function ModificaDesdeFormulario1(ByRef formulario As Form, Opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario1
    ModificaDesdeFormulario1 = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        ElseIf TypeOf Control Is TextBox And Control.visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or (Opcion = 3 And Control.Name = "txtAux") Then
            If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                 If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                                 cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
    
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox And Control.visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        ElseIf TypeOf Control Is OptionButton And Control.visible Then
            If Control.Enabled Then
                If Control.Value = True And Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        Aux = Control.Index
                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    conn.Execute Aux, , adCmdText

    ModificaDesdeFormulario1 = True
    Exit Function
    
EModificaDesdeFormulario1:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function

Public Sub FormateaCampo(vTex As TextBox)
    Dim mTag As CTag
    Dim cad As String
    On Error GoTo EFormateaCampo
    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        If vTex.Text <> "" Then
            If mTag.Formato <> "" Then
                cad = TransformaPuntosComas(vTex.Text)
                cad = Format(cad, mTag.Formato)
                vTex.Text = cad
            End If
        End If
    End If
EFormateaCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Sub


Public Function FormatoCampo(ByRef vTex As TextBox) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim cad As String
    
    On Error GoTo EFormatoCampo

    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        FormatoCampo = mTag.Formato
    End If
    
EFormatoCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


'Añade: CESAR
'Para utilizalo en el arreglaGrid
Public Function FormatoCampo2(ByRef objec As Object) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim cad As String

    On Error GoTo EFormatoCampo2

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        FormatoCampo2 = mTag.Formato
    End If
    
EFormatoCampo2:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


Public Function TipoCamp(ByRef objec As Object) As String
Dim mTag As CTag
Dim cad As String

    On Error GoTo ETipoCamp

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        TipoCamp = mTag.TipoDato
    End If

ETipoCamp:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef CADENA As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim Cont As Integer
Dim cad As String

    I = 0
    Cont = 1
    cad = ""
    Do
        J = I + 1
        I = InStr(J, CADENA, "|")
        If I > 0 Then
            If Cont = Orden Then
                cad = Mid(CADENA, J, I - J)
                I = Len(CADENA) 'Para salir del bucle
                Else
                    Cont = Cont + 1
            End If
        End If
    Loop Until I = 0
    RecuperaValor = cad
End Function


'-----------------------------------------------------------------------
'Deshabilitar ciertas opciones del menu
'EN funcion del nivel de usuario
'Esto es a nivel general, cuando el Toolba es el mismo

'Para ello en el tag del button tendremos k poner un numero k nos diara hasta k nivel esta permitido

Public Sub PonerOpcionesMenuGeneral(ByRef formulario As Form)
Dim I As Integer
Dim J As Integer
'Dim bol As Boolean

On Error GoTo EPonerOpcionesMenuGeneral
'bol = vSesion.Nivel < 2

'Añadir, modificar y borrar deshabilitados si no nivel
With formulario
    For I = 1 To .Toolbar1.Buttons.Count
        If .Toolbar1.Buttons(I).Tag <> "" Then
            J = Val(.Toolbar1.Buttons(I).Tag)
            If J < vUsu.Nivel Then
                .Toolbar1.Buttons(I).Enabled = False
            End If
        End If
    Next I
End With

Exit Sub
EPonerOpcionesMenuGeneral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub


Public Sub PonerModoMenuGral(ByRef formulario As Form, activo As Boolean)
Dim I As Integer
'Dim j As Integer

On Error GoTo PonerModoMenuGral

'Añadir, modificar y borrar deshabilitados si no Modo
    With formulario
        For I = 1 To .Toolbar1.Buttons.Count
            Select Case .Toolbar1.Buttons(I).ToolTipText
                Case "Nuevo"
                    .Toolbar1.Buttons(I).visible = Not .DeConsulta
                Case "Modificar", "Eliminar", "Imprimir"
                    .Toolbar1.Buttons(I).visible = Not .DeConsulta
                    .Toolbar1.Buttons(I).Enabled = activo
'                Case "Modificar"
'                Case "Eliminar"
'                Case "Imprimir"
            End Select
        Next I
        
        
        'El menu Visible
        .mnModificar.visible = Not .DeConsulta
        .mnEliminar.visible = Not .DeConsulta
        'El menu activo
        .mnModificar.Enabled = activo
        .mnEliminar.Enabled = activo
    End With
    
    
Exit Sub
PonerModoMenuGral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub

Public Sub PonerOpcionesMenuGeneralNew(formulario As Form)
Dim Control As Object
Dim I As Integer
Dim J As Integer
'Dim bol As Boolean

On Error GoTo EPonerOpcionesMenuGeneralNew
'bol = vSesion.Nivel < 2
'Añadir, modificar y borrar deshabilitados si no nivel
    For Each Control In formulario.Controls
'        Debug.Print Control.Name
        
        If Mid(Control.Name, 1, 2) = "mn" And Mid(Control.Name, 1, 7) <> "mnBarra" _
           And Control.Name <> "mnOpciones" Then
            J = Val(Control.HelpContextID)
            If J < vUsu.Nivel And J <> 0 Then
                Control.Enabled = False
            End If
        End If
    Next Control

Exit Sub
EPonerOpcionesMenuGeneralNew:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub



'Este modifica las claves prinipales y todo
'la sentenca del WHERE cod=1 and .. viene en claves
Public Function ModificaDesdeFormularioClaves(ByRef formulario As Form, Claves As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String
Dim I As Integer

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormularioClaves = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    cadWHERE = Claves
    'Construimos el SQL
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    conn.Execute Aux, , adCmdText

ModificaDesdeFormularioClaves = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function

Public Function BLOQUEADesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQUEADesdeFormulario
    BLOQUEADesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.visible = True Then
            If Control.Tag <> "" Then

                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        'Lo pondremos para el WHERE
                         If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                         cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                    End If
                End If
            End If
        End If
    Next Control

    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.tabla
        Aux = Aux & " WHERE " & cadWHERE & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario = True
    End If
    
EBLOQUEADesdeFormulario:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function


'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function BLOQUEADesdeFormulario2(ByRef formulario As Form, ByRef ado As Adodc, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim AntiguoCursor As Byte
Dim nomcamp As String

    On Error GoTo EBLOQUEADesdeFormulario2
    
    BLOQUEADesdeFormulario2 = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If (TypeOf Control Is TextBox) Or (TypeOf Control Is ComboBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Sea para el where o para el update esto lo necesito
                        'Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            Aux = ValorParaSQL(CStr(ado.Recordset.Fields(mTag.columna)), mTag)
                            'Lo pondremos para el WHERE
                             If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                             cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                        End If
                    End If
                End If
            End If
        End If
    Next Control

    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.tabla
        Aux = Aux & " WHERE " & cadWHERE & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario2 = True
    End If
    
EBLOQUEADesdeFormulario2:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla 2"
'        BLOQUEADesdeFormulario2 = False
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function BloqueaRegistro(cadTABLA As String, cadWHERE As String) As Boolean
Dim Aux As String

    On Error GoTo EBloqueaRegistro
        
    BloqueaRegistro = False
    Aux = "select * FROM " & cadTABLA
    Aux = Aux & " WHERE " & cadWHERE & " FOR UPDATE"

    'Intenteamos bloquear
    PreparaBloquear
    conn.Execute Aux, , adCmdText
    BloqueaRegistro = True

EBloqueaRegistro:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
End Function


Public Function BloqueaRegistroForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim AuxDef As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQ
    BloqueaRegistroForm = False
    Set mTag = New CTag
    Aux = ""
    AuxDef = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        Aux = ValorParaSQL(Control.Text, mTag)
                        AuxDef = AuxDef & Aux & "|"
                    End If
                End If
            End If
        End If
    Next Control

    If AuxDef = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
'        Aux = "Insert into zBloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & mTag.tabla
        Aux = Aux & "',""" & AuxDef & """)"
        conn.Execute Aux
        BloqueaRegistroForm = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If conn.Errors.Count > 0 Then
            If conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueaRegistroForm(ByRef TextBoxConTag As TextBox) As Boolean
Dim mTag As CTag
Dim Sql As String

'Solo me interesa la tabla
On Error Resume Next
    Set mTag = New CTag
    mTag.Cargar TextBoxConTag
    If mTag.Cargado Then
'        SQL = "DELETE from zBloqueos where codusu=" & vUsu.Codigo & " and tabla='" & mTag.tabla & "'"
        conn.Execute Sql
        If Err.Number <> 0 Then
            Err.Clear
        End If
    End If
    Set mTag = Nothing
End Function


'====================== LAURA

Public Function ComprobarCero(Valor As String) As String
    If Valor = "" Then
        ComprobarCero = "0"
    Else
        ComprobarCero = Valor
    End If
End Function

Public Sub InsertarCambios(tabla As String, ValorAnterior As String, numalbar As String)
Dim Sql As String
Dim Sql2 As String

    Sql = CadenaCambio

    Sql2 = "insert into cambios (codusu, fechacambio, tabla, numalbar, cadena, valoranterior) values ("
    Sql2 = Sql2 & DBSet(vSesion.Codusu, "N") & "," & DBSet(Now, "FH") & "," & DBSet(tabla, "T") & ","
    Sql2 = Sql2 & DBSet(numalbar, "T") & ","
    Sql2 = Sql2 & DBSet(Sql, "T") & ","
    If ValorAnterior = ValorNulo Then
        Sql2 = Sql2 & ValorNulo & ")"
    Else
        Sql2 = Sql2 & DBSet(ValorAnterior, "T") & ")"
    End If

    conn.Execute Sql2

End Sub
    
Public Sub CargarValoresAnteriores(formulario As Form, Optional opcio As Integer, Optional nom_frame As String)
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim cad As String
    Set mTag = New CTag

    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            If Izda <> "" Then Izda = Izda & " , "
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.columna & " = "
                            'Parte VALUES
                            cad = ValorParaSQL(Control.Text, mTag)
                            Izda = Izda & cad
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Izda <> "" Then Izda = Izda & " , "
                    'Access
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & " = "
                    If Control.Value = 1 Then
                        cad = "1"
                        Else
                        cad = "0"
                    End If
                    If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                    Izda = Izda & cad
                End If
            End If
            
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Izda <> "" Then Izda = Izda & " , "
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & " = "
                        If Control.ListIndex = -1 Then
                            cad = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            cad = Control.ItemData(Control.ListIndex)
                        Else
                            cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        Izda = Izda & cad
                    End If
                End If
            End If
            
        'OPTION BUTTON
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            If Izda <> "" Then Izda = Izda & " , "
                            Izda = Izda & "" & mTag.columna & " = "
                            cad = Control.Index
                            Izda = Izda & cad
                        End If
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is DTPicker Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Izda <> "" Then Izda = Izda & " , "
                        Izda = Izda & "" & mTag.columna & " = "
                        
                        'Parte VALUES
                        If Control.visible Then
                            cad = ValorParaSQL(Control.Value, mTag)
                        Else
                            cad = ValorNulo
                        End If
                        Izda = Izda & cad
                    End If
                End If
            End If
        End If
        
    Next Control

    ValorAnterior = Izda

End Sub


Public Sub CalcularImporteNue(ByRef Cantidad As TextBox, ByRef Precio As TextBox, ByRef Importe As TextBox, Tipo As Integer)
'Calcula el Importe de una linea de hcode facturas
Dim vImp As Currency
Dim vCan As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    Cantidad = ComprobarCero(Cantidad.Text)
    Precio = ComprobarCero(Precio.Text)
    Importe = ComprobarCero(Importe.Text)
    
    Select Case Tipo
        Case 0 ' me han introducido la cantidad
            vImp = CCur(ImporteFormateado(Cantidad.Text)) * CCur(ImporteFormateado(Precio.Text))
            vImp = Round2(vImp, 2)
            Importe.Text = Format(vImp, "###,##0.00")
        Case 1 ' me han introducido el precio
            vImp = CCur(ImporteFormateado(Cantidad.Text)) * CCur(ImporteFormateado(Precio.Text))
            vImp = Round2(vImp, 2)
            Importe.Text = Format(vImp, "###,##0.00")
        Case 2 ' me han introducido el importe
            vCan = CCur(ImporteFormateado(Importe.Text)) / CCur(ImporteFormateado(Precio.Text))
            vCan = Round2(vCan, 3)
            Cantidad.Text = Format(vCan, "##,##0.000")
    End Select
    
End Sub


'Public Function PonerNomEmple(codEmp As String) As String
'Dim nomEmp As String
'Dim cad As String
'
'    'apellidos i nombre del empleado
'    If (codEmp <> "") Then
'        nomEmp = "nomemple"
'        cad = DevuelveDesdeBDNew(cAgro, "empleado", "apeemple", "codemple", codEmp, "N", nomEmp, "codempre", CStr(vSesion.Empresa), "N", "codagenc", CStr(vSesion.Agencia), "N")
'        If cad <> "" Then cad = cad & ", " & nomEmp
'    End If
'    PonerNomEmple = cad
'End Function



Public Function ExisteCP(T As TextBox) As Boolean
'comprueba para un campo de texto que sea clave primaria, si ya existe un
'registro con ese valor
Dim vtag As CTag
Dim devuelve As String

    On Error GoTo ErrExiste

    ExisteCP = False
    If T.Text <> "" Then
        If T.Tag <> "" Then
            Set vtag = New CTag
            If vtag.Cargar(T) Then
'                If vtag.EsClave Then
                    devuelve = DevuelveDesdeBDNew(cAgro, vtag.tabla, vtag.columna, vtag.columna, T.Text, vtag.TipoDato)
                    If devuelve <> "" Then
    '                    MsgBox "Ya existe un registro para " & vtag.Nombre & ": " & T.Text, vbExclamation
                        MsgBox "Ya existe el " & vtag.Nombre & ": " & T.Text, vbExclamation
                        ExisteCP = True
                        PonerFoco T
                    End If
'                End If
            End If
            Set vtag = Nothing
        End If
    End If
    Exit Function
    
ErrExiste:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar código.", Err.Description
End Function




Public Function TotalRegistros(vSQL As String) As Long
'Devuelve el valor de la SQL
'para obtener COUNT(*) de la tabla
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalRegistros = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then TotalRegistros = Rs.Fields(0).Value  'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        TotalRegistros = 0
        Err.Clear
    End If
End Function



Public Function Round2(Number As Variant, Optional NumDigitsAfterDecimals As Long) As Variant
Dim Ent As Integer
Dim cad As String

  ' Comprobaciones

  If Not IsNumeric(Number) Then
    Err.Raise 13, "Round2", "Error de tipo. Ha de ser un número."
    Exit Function
  End If

  If NumDigitsAfterDecimals < 0 Then
    Err.Raise 0, "Round2", "NumDigitsAfterDecimals no puede ser negativo."
    Exit Function
  End If

  ' Redondeo.

  cad = "0"
  If NumDigitsAfterDecimals <> 0 Then cad = cad & "." & String(NumDigitsAfterDecimals, "0")
  Round2 = Val(TransformaComasPuntos(Format(Number, cad)))

End Function

Private Function RellenaABlancos(CADENA As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim cad As String
    
    cad = Space(Longitud)
    If PorLaDerecha Then
        cad = CADENA & cad
        RellenaABlancos = Left(cad, Longitud)
    Else
        cad = cad & CADENA
        RellenaABlancos = Right(cad, Longitud)
    End If
    
End Function

Public Function ObtenerLetraSerie(tipMov As String) As String
'Devuelve la letra de serie asociada al tipo de movimiento
Dim LEtra As String
Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error Resume Next
    
    Sql = "select letraser from usuarios.stipom where codtipom = " & DBSet(tipMov, "T")
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    LEtra = ""
    If Not Rs.EOF Then
        LEtra = DBLet(Rs.Fields(0).Value, "T")
    End If
'--monica: cambiado por el recordset anterior pq stipom está en la bd usuarios
'    LEtra = DevuelveDesdeBDNew(cAgro, "stipom", "letraser", "codtipom", tipMov, "T")
    If LEtra = "" Then MsgBox "Las factura de venta no tienen asignada una letra de serie", vbInformation
    ObtenerLetraSerie = LEtra
End Function

Public Function CalcularPorcentaje(Importe As Currency, Porce As Currency, NumDecimales As Long) As Variant
'devuelve el valor del Porcentaje aplicado al Importe
'Ej el 16% de 120 = 19.2
'Dim vImp As Currency
'Dim vDto As Currency
    
    On Error Resume Next
'
'    Importe = ComprobarCero(Importe)
'    Dto = ComprobarCero(Dto)
'
'    vImp = CCur(Importe)
'    vDto = CCur(Dto)
    
    
    'vImp = Round(vImp, 2)
    
    CalcularPorcentaje = Round2((Importe * Porce) / 100, NumDecimales)
    
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function CalcularDto(Importe As String, Dto As String) As String
'devuelve el Dto% del Importe
'Ej el 16% de 120 = 19.2
Dim vImp As Currency
Dim vDto As Currency
On Error Resume Next

    Importe = ComprobarCero(Importe)
    Dto = ComprobarCero(Dto)
    
    vImp = CCur(Importe)
    vDto = CCur(Dto)
    
    vImp = ((vImp * vDto) / 100)
    'vImp = Round(vImp, 2)
    
    CalcularDto = CStr(vImp)
    If Err.Number <> 0 Then Err.Clear
End Function

'---------------------------------------------------------------------------------
'
'       Para buscar en los checks con las dos opciones de true y false
'
'A partir de un check cualquiera devolvera nombre e indice, si tiene. Si no sera ()
Public Sub CheckBusqueda(ByRef CH As CheckBox)
    NombreCheck = ""
    NombreCheck = CH.Name & "("
    On Error Resume Next
    NombreCheck = NombreCheck & CH.Index
    If Err.Number <> 0 Then Err.Clear
    NombreCheck = NombreCheck & ")"
End Sub

Public Sub CheckCadenaBusqueda(ByRef CH As CheckBox, ByRef CadenaCHECKs As String)
        CheckBusqueda CH
        If InStr(1, CadenaCHECKs, NombreCheck) = 0 Then CadenaCHECKs = CadenaCHECKs & NombreCheck & "|"
End Sub

Public Function PonerAlmacen(codAlm As String) As String
'Comprueba si existe el Almacen y lo pone en el Text
Dim devuelve As String
    
    On Error Resume Next

    If codAlm = "" Then
        MsgBox "Debe introducir el Almacen.", vbInformation
    Else
        devuelve = DevuelveDesdeBDNew(cAgro, "salmpr", "codalmac", "codalmac", codAlm, "N")
        If devuelve = "" Then
            MsgBox "No existe el Almacen: " & Format(codAlm, "000"), vbInformation
            PonerAlmacen = ""
        Else
            PonerAlmacen = Format(codAlm, "000")
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function CalcularImporteFClien(Cantidad As String, Precio As String, Dto1 As String, Dto2 As String, TipoDto As Byte, ImpDto As String, Optional Bruto As String) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
'Si DtoProv=sprove.tipodtos, calcular Importe para Proveedores y obtener el tipo de descuento
'del campo sprove.tipodtos, si es para Clientes obtener el tipo de descuento del
'parametro spara1.tipodtos
'Tipo Descuento: 0=aditivo, 1=sobre resto
Dim vImp As Currency
Dim vDto1 As Currency, vDto2 As Currency
Dim vPre As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    Cantidad = ComprobarCero(Cantidad)
    vPre = ComprobarCero(Precio)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
'    If vParamAplic.Cooperativa <> 5 And vParamAplic.Cooperativa <> 4 Then ' castelduc calcula primero comisiones y luego el importe
'[Monica]16/09/2011: modificado con el parametro
    If vParamAplic.TipoCalculoComision = 0 Then
        If Bruto <> "" Then
            vImp = CCur(Bruto) - CCur(ImpDto)
        Else
            vImp = (CCur(Cantidad) * CCur(vPre)) - CCur(ImpDto)
        End If
    Else
        If Bruto <> "" Then
            vImp = CCur(Bruto)
        Else
            vImp = (CCur(Cantidad) * CCur(vPre))
        End If
    End If
        
    If TipoDto = 0 Then 'Dto Aditivo
        vDto1 = (CCur(Dto1) * vImp) / 100
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto1 - vDto2
    ElseIf TipoDto = 1 Then 'Sobre Resto
        vDto1 = (CCur(Dto1) * vImp) / 100
        vImp = vImp - vDto1
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto2
    End If

'[Monica]16/09/2011: modificado con el parametro
'    If vParamAplic.Cooperativa = 5 Or vParamAplic.Cooperativa = 4 Then
    If vParamAplic.TipoCalculoComision = 1 Then
        vImp = vImp - CCur(ImpDto)
    End If
    
    vImp = Round(vImp, 2)
    CalcularImporteFClien = CStr(vImp)
    
End Function


Public Function CalcularImporte(Cantidad As String, Precio As String, Dto1 As String, Dto2 As String, TipoDto As Byte, ImpDto As String, Optional Bruto As String) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
'Si DtoProv=sprove.tipodtos, calcular Importe para Proveedores y obtener el tipo de descuento
'del campo sprove.tipodtos, si es para Clientes obtener el tipo de descuento del
'parametro spara1.tipodtos
'Tipo Descuento: 0=aditivo, 1=sobre resto
Dim vImp As Currency
Dim vDto1 As Currency, vDto2 As Currency
Dim vPre As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    Cantidad = ComprobarCero(Cantidad)
    vPre = ComprobarCero(Precio)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
    If Bruto <> "" Then
        vImp = CCur(Bruto) - CCur(ImpDto)
    Else
        vImp = (CCur(Cantidad) * CCur(vPre)) - CCur(ImpDto)
    End If
        
    If TipoDto = 0 Then 'Dto Aditivo
        vDto1 = (CCur(Dto1) * vImp) / 100
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto1 - vDto2
    ElseIf TipoDto = 1 Then 'Sobre Resto
        vDto1 = (CCur(Dto1) * vImp) / 100
        vImp = vImp - vDto1
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto2
    End If
    
    vImp = Round(vImp, 2)
    CalcularImporte = CStr(vImp)
End Function






Public Function EsProveedorVarios(codProve As String) As Boolean
Dim devuelve As String

    EsProveedorVarios = False
    devuelve = DevuelveDesdeBD("provario", "proveedor", "codprove", codProve, "N")
    If devuelve <> "" Then EsProveedorVarios = CBool(devuelve)
    'Es proveedor de varios Y podemos recuperar de ????
End Function

Public Function QuitarCero(Valor As String) As String
    On Error Resume Next
    
    If Valor <> "" Then
        If CSng(Valor) = 0 Then
            QuitarCero = ""
        Else
            QuitarCero = Valor
        End If
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function BloqueoManual(cadTABLA As String, cadWHERE As String)
Dim Aux As String

On Error GoTo EBLOQ

    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "INSERT INTO zbloqueos(codusu,tabla,clave) VALUES(" & vUsu.codigo & ",'" & cadTABLA
        Aux = Aux & "',""" & cadWHERE & """)"
        conn.Execute Aux
        BloqueoManual = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If conn.Errors.Count > 0 Then
            If conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
'    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueoManual(cadTABLA As String) As Boolean
Dim Sql As String

'Solo me interesa la tabla
On Error Resume Next

        Sql = "DELETE FROM zbloqueos WHERE codusu=" & vUsu.codigo & " and tabla='" & cadTABLA & "'"
        conn.Execute Sql
        If Err.Number <> 0 Then
            Err.Clear
        End If
End Function

'++monica
'funcion de la libreria general de gessocial de Rafa, necesaria para pasar al aridoc
Public Function CApos(Texto As String) As String
    Dim I As Integer
    I = InStr(1, Texto, "'")
    If I = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, I) & "'" & Mid(Texto, I + 1, Len(Texto) - I)
    End If
    '-- Ya que estamos transformamos las Ñ
    Texto = CApos
    I = InStr(1, Texto, "¥")
    If I = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, I - 1) & "Ñ" & Mid(Texto, I + 1, Len(Texto) - I)
    End If
    '-- Y otra más
    Texto = CApos
    I = InStr(1, Texto, "¾")
    If I = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, I - 1) & "Ñ" & Mid(Texto, I + 1, Len(Texto) - I)
    End If
    '-- Seguimos con transformaciones
    Texto = CApos
    I = InStr(1, Texto, "¦")
    If I = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, I - 1) & "ª" & Mid(Texto, I + 1, Len(Texto) - I)
    End If
End Function



Public Function DevuelveValor(vSQL As String) As Variant
'Devuelve el valor de la SQL
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    DevuelveValor = 0
    If Not Rs.EOF Then
        'antes RS.Fields(0).Value > 0
        If Not IsNull(Rs.Fields(0).Value) Then DevuelveValor = Rs.Fields(0).Value   'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        DevuelveValor = 0
        Err.Clear
    End If
End Function


Public Function DevuelveValorConta(vSQL As String) As Variant
'Devuelve el valor de la SQL
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, ConnConta, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    DevuelveValorConta = 0
    If Not Rs.EOF Then
        'antes RS.Fields(0).Value > 0
        If Not IsNull(Rs.Fields(0).Value) Then DevuelveValorConta = Rs.Fields(0).Value   'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        DevuelveValorConta = 0
        Err.Clear
    End If
End Function





Public Function TotalRegistrosConsulta(cadSQL) As Long
Dim cad As String
Dim Rs As ADODB.Recordset

    On Error GoTo ErrTotReg
    cad = "SELECT count(*) FROM (" & cadSQL & ") x"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not Rs.EOF Then
        TotalRegistrosConsulta = DBLet(Rs.Fields(0).Value, "N")
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
ErrTotReg:
    MuestraError Err.Number, "", Err.Description
End Function


Public Sub BorrarArchivo(nomFich As String)
    
    On Error Resume Next
    
    If Dir(nomFich) <> "" Then Kill nomFich

    If Err.Number <> 0 Then MuestraError Err.Number, "Borrar fichero"

End Sub

'++monica: facturamos segun el campo de forfaits
Public Function TipoFacturarForfaits(Albaran As String, Linea As String) As Byte
' devuelve 0: facturar por unidades
'          1: facturar por kilos
Dim Rs As ADODB.Recordset
Dim Sql As String

    TipoFacturarForfaits = 2
    
    If Trim(Albaran) = "" Or Trim(Linea) = "" Then Exit Function

    Sql = "select forfaits.facturar from albaran_variedad, forfaits "
    Sql = Sql & " where albaran_variedad.numalbar = " & DBSet(Albaran, "N")
    Sql = Sql & " and albaran_variedad.numlinea = " & DBSet(Linea, "N")
    Sql = Sql & " and forfaits.codforfait = albaran_variedad.codforfait "
    Sql = Sql & " order by numlinea "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If Not Rs.EOF Then
        TipoFacturarForfaits = DBLet(Rs.Fields(0).Value, "N")
    End If
    
End Function


Public Function ExisteTabla(tabla As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo eExisteTabla
    
    ExisteTabla = False
    
    Sql = "describe " & tabla
    conn.Execute Sql

eExisteTabla:
    ExisteTabla = (Err.Number = 0)
End Function


Public Function EsArticuloRetornable(codigo As String) As Boolean
Dim Sql As String

    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "stipar", "esretornable", "codtipar", codigo, "T")

    EsArticuloRetornable = (Sql = "1")

End Function


Public Function EsTransportista(codigo As String) As Boolean
Dim Sql As String

    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "agencias", "tipo", "codtrans", codigo, "N")

    EsTransportista = (Sql = "0")

End Function


Public Function NroAlbaranAsignado(PaletPedido As String, Tipo As Boolean) As String
'Tipo 0 = palet
'     1 = pedido
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo eNroAlbaranAsignado

    NroAlbaranAsignado = ""
    
    Select Case Tipo
        Case 0 'palet
            Sql = "select albaran.numalbar from (palets INNER JOIN pedidos ON palets.numpedid = pedidos.numpedid) "
            Sql = Sql & " INNER JOIN albaran ON  pedidos.numalbar = albaran.numalbar "
            Sql = Sql & " where palets.numpalet = " & DBSet(PaletPedido, "N")
        Case 1 'pedido
            Sql = "select pedidos.numalbar from pedidos where numpedid = " & DBSet(PaletPedido, "N")
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        NroAlbaranAsignado = DBLet(Rs.Fields(0).Value, "T")
    End If
    Exit Function

eNroAlbaranAsignado:
    MuestraError Err.Number, "Numero de Albaran Asignado", Err.Description
End Function

Public Sub PonerContRegIndicador(ByRef lblIndicador As label, ByRef vData As Adodc, cadBuscar As String)
'cuando esta en el MODO 2 pone el label de contador de registros añadiendo
'la palabra 'Busqueda' si es el resultado de una busqueda
'devolvera: "1 de 20" o "BUSQUEDA: 1 de 20"
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    cadReg = PonerContRegistros(vData) 'devuelve: "1 de 20"
    
    If cadBuscar = "" Or cadReg = "" Then
        lblIndicador.Caption = cadReg
    Else
        lblIndicador.Caption = "BUSQUEDA: " & cadReg
    End If
End Sub

Public Sub AyudaSocios(frmBas As frmBasico, Optional CodActual As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|800|;S|txtAux(1)|T|Descripción|3800|;"
    frmBas.CadenaConsulta = "SELECT rsocios.codsocio, rsocios.nomsocio "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM rsocios "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "Código|N|N|0|999999|rsocios|codsocio|000000|S|"
    frmBas.Tag2 = "Descripción|T|N|||rsocios|nomsocio|||"
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 40
    frmBas.DeConsulta = True
    frmBas.tabla = "rsocios"
    frmBas.CampoCP = "codsocio"
    frmBas.Report = "rManSocios.rpt"
    frmBas.Caption = "Socios"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.Show vbModal

End Sub


Public Sub AyudaZonasCC(frmBas As frmBasico, Optional CodActual As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|800|;S|txtAux(1)|T|Descripción|3930|;"
    frmBas.CadenaConsulta = "SELECT cczonas.codzona, cczonas.nomzona "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM cczonas "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "Código|N|N|0|9999|cczonas|codzona|0000|S|"
    frmBas.Tag2 = "Descripción|T|N|||cczonas|nomzona|||"
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 30
    frmBas.DeConsulta = True
    frmBas.tabla = "cczonas"
    frmBas.CampoCP = "codzona"
    frmBas.Report = "rManCCZonas.rpt"
    frmBas.Caption = "Zonas"
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaAreasCC(frmBas As frmBasico, Optional CodActual As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|800|;S|txtAux(1)|T|Descripción|3930|;"
    frmBas.CadenaConsulta = "SELECT ccareas.codarea, ccareas.nomarea "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM ccareas "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "Código|N|N|0|9999|ccareas|codarea|0000|S|"
    frmBas.Tag2 = "Descripción|T|N|||ccareas|nomarea|||"
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 50
    frmBas.tabla = "ccareas"
    frmBas.CampoCP = "codarea"
    frmBas.Report = "rManCCAreas.rpt"
    frmBas.Caption = "Áreas"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaTrabajadores(frmBas As frmBasico, Optional CodActual As String)
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|800|;S|txtAux(1)|T|Descripción|3800|;"
    frmBas.CadenaConsulta = "SELECT straba.codtraba, straba.nomtraba "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM straba "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "Código|N|N|0|999999|straba|codtraba|000000|S|"
    frmBas.Tag2 = "Descripción|T|N|||straba|nomtraba|||"
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 40
    frmBas.DeConsulta = True
    frmBas.tabla = "straba"
    frmBas.CampoCP = "codtraba"
    frmBas.Report = "rManTraba.rpt"
    frmBas.Caption = "Trabajadores"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.Show vbModal
End Sub


Public Sub AyudaCategorias(frmBas As frmBasico, Optional CodActual As String)
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|600|;S|txtAux(1)|T|Descripción|4000|;"
    frmBas.CadenaConsulta = "SELECT salarios.codcateg, salarios.nomcateg "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM salarios "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "Código|N|N|0|99|salarios|codcateg|00|S|"
    frmBas.Tag2 = "Descripción|T|N|||salarios|nomcateg|||"
    frmBas.Maxlen1 = 2
    frmBas.Maxlen2 = 44
    frmBas.DeConsulta = True
    frmBas.tabla = "salarios"
    frmBas.CampoCP = "codcateg"
    frmBas.Report = "rManTraba.rpt"
    frmBas.Caption = "Salarios"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.Show vbModal
End Sub

' Funcion que inserta o modifica las lineas de FACTURAS_CALIBRE
' se usa en el mantenimiento de facturas de cliente y en el paso de albaranes a facturas

Public Function InsertarModificarCalibres(Insertar As Boolean, codTipoM As String, Factura As String, FecFactu As String, NumLinea As String, Albaran As String, NumlineaAlb As String, TCantReal As String, TUnidades As String, TImpBruto As String, TImpNeto As String, MenError As String) As Boolean
' Insertar : = true : insertamos todas las lineas en facturas_calibre del albaran prorrateando
'            = false: venimos de modificar lineas en facturas_variedad prorrateamos lineas de facturas_calibre segun los cambios que hay en facturas_variedad
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String
Dim vImpDto As Currency
Dim vDto1 As Currency
Dim vDto2 As Currency
Dim vImpNeto As Currency
Dim vImpBruto As Currency
Dim vPrecNeto As Currency
Dim vPrecBruto As Currency

Dim TipoDto As String
Dim ImpDto As String
Dim Cliente As String
Dim Rdo As Long

Dim ImpBrutoAc As Currency
Dim ImpNetoAc As Currency

Dim Diferencia As Currency
Dim Diferencia1 As Currency

Dim UltimaLinea As Currency
Dim TipoFactFor As Byte

Dim vHayReg As Byte
Dim KilosCaja As Currency

    On Error GoTo eInsertarModificarCalibres


    KilosCaja = DevuelveValor("select kiloscaj from forfaits inner join albaran_variedad on forfaits.codforfait = albaran_variedad.codforfait where albaran_variedad.numalbar = " & DBSet(Albaran, "N") & " and numlinea = " & DBSet(NumlineaAlb, "N"))

    ' Si venimos de insertar una linea de factura, insertamos automaticamente todas las lineas de calibre prorrateando
    If Insertar Then
        ' Primero insertamos con los precios e importes a cero
        
        Sql = "insert into facturas_calibre (codtipom,numfactu,fecfactu,numlinea,numline1,numalbar,numlinealbar,numline1albar,cantreal,cantfact,"
        Sql = Sql & " precibru,precinet,dtocom1,dtocom2,imporbru,impornet,unidades) "
        Sql = Sql & " select " & DBSet(codTipoM, "T") & "," & DBSet(Factura, "N") & "," & DBSet(FecFactu, "F") & ","
        Sql = Sql & DBSet(NumLinea, "N") & ",numline1," & DBSet(Albaran, "N") & "," & DBSet(NumlineaAlb, "N") & ",numline1,"
        Sql = Sql & " pesoneto, round(numcajas * " & DBSet(KilosCaja, "N") & ",2), 0,0,0,0,0,0,unidades "
        Sql = Sql & " from albaran_calibre where numalbar = " & DBSet(Albaran, "N")
        Sql = Sql & " and numlinea = " & DBSet(NumlineaAlb, "N")
        Sql = Sql & " order by numline1 "
        
        conn.Execute Sql
    End If
    
    ' Prorrateamos TODO con respecto a los kilos
    Sql = "select * from facturas_calibre where codtipom = " & DBSet(codTipoM, "T") & " and numfactu = " & DBSet(Factura, "N")
    Sql = Sql & " and fecfactu = " & DBSet(FecFactu, "F") & " and numlinea = " & DBSet(NumLinea, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "facturas", "impdtoc", "codtipom", codTipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vImpDto = ComprobarCero(Sql)
    
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "facturas", "dtocom1", "codtipom", codTipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vDto1 = ComprobarCero(Sql)
    
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "facturas", "dtocom2", "codtipom", codTipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vDto2 = ComprobarCero(Sql)
    
    '++monica:030608:traemos el redondeo del precio
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "facturas", "codclien", "codtipom", codTipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    Cliente = ComprobarCero(Sql)
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "clientes", "nrodecprec", "codclien", Cliente, "N")
    Rdo = ComprobarCero(Sql)
    
    vHayReg = 0
    
    
    ImpBrutoAc = 0
    ImpNetoAc = 0
    
    While Not Rs.EOF
        vHayReg = 1
        
        TipoDto = DevuelveDesdeBDNew(cAgro, "clientes", "tipodtos", "codclien", Cliente, "N")
        If TipoFacturarForfaits(CStr(Albaran), CStr(NumlineaAlb)) = 1 Then 'kilos
            TipoFactFor = 1
            
            vImpBruto = Round2(TImpBruto * DBLet(Rs!cantreal, "N") / TCantReal, 2)
            ImpBrutoAc = ImpBrutoAc + vImpBruto
            
            vImpNeto = Round2(TImpNeto * DBLet(Rs!cantreal, "N") / TCantReal, 2)
            ImpNetoAc = ImpNetoAc + vImpNeto
                
            '[Monica]24/11/2011: si las unidades son 0 no hay division
            'precio neto
            vPrecNeto = 0
            vPrecBruto = 0
            If DBLet(Rs!cantreal, "N") <> 0 Then
                vPrecNeto = Round2(vImpNeto / DBLet(Rs!cantreal, "N"), Rdo)
                vPrecBruto = Round2(vImpBruto / DBLet(Rs!cantreal, "N"), Rdo)
            End If
            '++monica:040608 : solo si redondeo <> 4
            If Rdo = 2 Or Rdo = 3 Then
                vImpNeto = Round2(vPrecNeto * DBLet(Rs!cantreal, "N"), 2)
            End If
            
        Else 'unidades
            TipoFactFor = 0
'            ImpDto = CalcularImporteDto(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), TipoM, Factura, FecFactu, CStr(vImpDto), True)
'            vImpNeto = CalcularImporteFClien(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), CStr(vDto1), CStr(vDto2), CByte(TipoDto), CStr(ImpDto), DBLet(Rs!imporbru, "N"))

            vImpBruto = 0
            If TUnidades <> 0 Then
                vImpBruto = Round2(TImpBruto * DBLet(Rs!Unidades, "N") / TUnidades, 2)
            End If
            ImpBrutoAc = ImpBrutoAc + vImpBruto
            
            vImpNeto = 0
            If TUnidades <> 0 Then
                vImpNeto = Round2(TImpNeto * DBLet(Rs!Unidades, "N") / TUnidades, 2)
            End If
            ImpNetoAc = ImpNetoAc + vImpNeto
            
            '[Monica]24/11/2011: si las unidades son 0 no hay division
            'precio neto
            vPrecNeto = 0
            vPrecBruto = 0
            If DBLet(Rs!Unidades, "N") <> 0 Then
                vPrecNeto = Round2(vImpNeto / DBLet(Rs!Unidades, "N"), Rdo)
                vPrecBruto = Round2(vImpBruto / DBLet(Rs!Unidades, "N"), Rdo)
            End If
            
            '++monica:040608
            If Rdo = 2 Or Rdo = 3 Then
                vImpNeto = Round2(vPrecNeto * DBLet(Rs!Unidades, "N"), 2)
                vImpBruto = Round2(vPrecBruto * DBLet(Rs!Unidades, "N"), 2)
            End If
            
        End If
        
        
        Sql2 = "update facturas_calibre set "
        Sql2 = Sql2 & "precibru = " & DBSet(vPrecBruto, "N")
        Sql2 = Sql2 & ",precinet = " & DBSet(vPrecNeto, "N")
        Sql2 = Sql2 & ",imporbru = " & DBSet(vImpBruto, "N")
        Sql2 = Sql2 & ",impornet = " & DBSet(vImpNeto, "N")
        Sql2 = Sql2 & ",dtocom1 = " & DBSet(vDto1, "N")
        Sql2 = Sql2 & ",dtocom2 = " & DBSet(vDto2, "N")
        Sql2 = Sql2 & " where codtipom = " & DBSet(codTipoM, "T")
        Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
        Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
        Sql2 = Sql2 & " and numlinea = " & DBSet(NumLinea, "N")
        Sql2 = Sql2 & " and numline1 = " & DBSet(Rs!numline1, "N")
    
        conn.Execute Sql2
    
        UltimaLinea = DBLet(Rs!numline1, "N")
    
        Rs.MoveNext
    Wend
    
    Rs.Close
    
    '[Monica]16/09/2011: si no coincide la suma con los totales redondeamos en la ultima linea
    If vHayReg = 1 Then
        If ImpBrutoAc <> TImpBruto Or ImpNetoAc <> TImpNeto Then
            Diferencia = TImpBruto - ImpBrutoAc
            Diferencia1 = TImpNeto - ImpNetoAc
            
            Sql2 = "update facturas_calibre set impornet = impornet + " & DBSet(Diferencia1, "N")
            Sql2 = Sql2 & ", imporbru = imporbru + " & DBSet(Diferencia, "N")
            Sql2 = Sql2 & " where codtipom = " & DBSet(codTipoM, "T")
            Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
            Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
            Sql2 = Sql2 & " and numlinea = " & DBSet(NumLinea, "N")
            Sql2 = Sql2 & " and numline1 = " & DBSet(UltimaLinea, "N")
        
            conn.Execute Sql2
    
            If TipoFactFor = 1 Then 'kilos
                Sql2 = "update facturas_calibre set precinet = round(impornet / cantreal, " & DBSet(Rdo, "N") & ") "
                Sql2 = Sql2 & " where codtipom = " & DBSet(codTipoM, "T")
                Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
                Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
                Sql2 = Sql2 & " and numlinea = " & DBSet(NumLinea, "N")
                Sql2 = Sql2 & " and numline1 = " & DBSet(UltimaLinea, "N")
            
                conn.Execute Sql2
            Else 'unidades
                'precio neto
                Sql2 = "update facturas_calibre set precinet = round(impornet / unidades, " & DBSet(Rdo, "N") & ") "
                Sql2 = Sql2 & " where codtipom = " & DBSet(codTipoM, "T")
                Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
                Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
                Sql2 = Sql2 & " and numlinea = " & DBSet(NumLinea, "N")
                Sql2 = Sql2 & " and numline1 = " & DBSet(UltimaLinea, "N")
            
                conn.Execute Sql2
            End If
        End If
    End If
    
    Set Rs = Nothing
    
    InsertarModificarCalibres = True
    Exit Function

eInsertarModificarCalibres:
    If Err.Number <> 0 Then
        MenError = MenError & vbCrLf & Err.Description
        InsertarModificarCalibres = False
    End If
End Function



