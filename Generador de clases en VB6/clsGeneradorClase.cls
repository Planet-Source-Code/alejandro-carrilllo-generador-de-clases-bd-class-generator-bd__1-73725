VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGeneradorClase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales para almacenar los valores de las propiedades
Private mvarTabla As String 'copia local
Private mvarSQLObtener As String 'copia local
Private mvarSchema As String 'copia local
Private mvarCadenaConn As String 'copia local
Private mvarMensajeError As String 'copia local
Private mvarRutaArchivoCLS As String 'copia local
Public Property Let RutaArchivoCLS(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.RutaArchivoCLS = 5
    mvarRutaArchivoCLS = vData
End Property


Public Property Get RutaArchivoCLS() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.RutaArchivoCLS
    RutaArchivoCLS = mvarRutaArchivoCLS
End Property


Public Function GenerarClase() As String
Dim strPlantilla As String, strPlantillaCol As String, strCadFunction As String
Dim strPrivates As String, strPropertys As String, strNombreClase As String
Dim strCompararCamposLlave As String, strCamposllaveaparametrosSelect As String, strCamposRecordset As String
Dim strCamposaparametrosSelect As String, strCamposaparametrosUpdate As String, strCamposaparametrosInsert As String
Dim conn As ADODB.Connection
Dim rsPrimaria As ADODB.Recordset, rsEstrucCol As ADODB.Recordset
Dim rsColumnas As ADODB.Recordset, rsSelect As ADODB.Recordset

On Local Error GoTo er:

  mvarMensajeError = ""
  'requisitos
  If mvarCadenaConn = "" Then
    mvarMensajeError = "La propiedad CadenaConn debe estar diligenciada para poder generar la clase"
    Exit Function
  End If
'  If mvarSchema = "" Then
'    mvarMensajeError = "La propiedad Schema debe estar diligenciada para poder generar la clase"
'    Exit Function
'  End If
  If mvarTabla = "" And mvarSQLObtener = "" Then
    mvarMensajeError = "La propiedad Tabla o la propiedad SQLObtener deben estar diligenciadas para poder generar la clase"
    Exit Function
  End If
  'inicia la función
  Set conn = New ADODB.Connection
  conn.CursorLocation = adUseClient
  conn.Open mvarCadenaConn
  If mvarTabla <> "" Then
      'busca las columnas y la llave de la tabla
    ''llave primaria
    Set rsPrimaria = conn.OpenSchema(adSchemaPrimaryKeys, Array(Empty, Empty, mvarTabla))
    Set rsPrimaria.ActiveConnection = Nothing
    '''la tabla debe tener por lo menos un campo como llave primaria
    If rsPrimaria.RecordCount < 1 Then
      mvarMensajeError = "La tabla debe tener una llave primaria"
      Exit Function
    End If
    ''Otras columnas
    '''estructura de otras columnas para campo null
    Set rsEstrucCol = conn.OpenSchema(adSchemaColumns, Array(Empty, Empty, mvarTabla, Empty))
    Set rsEstrucCol.ActiveConnection = Nothing
    '''otros datos de la estructura de otras columnas
    Set rsColumnas = conn.Execute("select * from " & mvarTabla & " where 1=2")
    Set rsColumnas.ActiveConnection = Nothing
  End If
  'recordset para select
  If mvarSQLObtener <> "" Then
    Set rsSelect = conn.Execute(mvarSQLObtener & " where 1=2")
    Set rsSelect.ActiveConnection = Nothing
  End If
  
  conn.Close
  'abre la plantilla de la clase para reemplazar las cadenas armadas
  If mvarTabla <> "" Then
    Open App.Path + "\plantillaclase.txt" For Input As #1
      strPlantilla = Input(LOF(1), 1)
    Close #1
  Else
    Open App.Path + "\plantillaSQL.txt" For Input As #1
      strPlantilla = Input(LOF(1), 1)
    Close #1
  End If
  'abre la plantilla de la coleccion para reemplazar las cadenas armadas
  Open App.Path + "\plantillacoleccion.txt" For Input As #1
    strPlantillaCol = Input(LOF(1), 1)
  Close #1
  'variables con las que se arma la clase
  strPrivates = ""
  strPropertys = ""
  strCompararCamposLlave = ""
  strCamposllaveaparametrosSelect = ""
  strCamposaparametrosSelect = ""
  strCamposaparametrosInsert = ""
  strCamposaparametrosUpdate = ""
  strCamposRecordset = ""
  strCamposParRecordset = ""
  strMapeoCamposRecordSetCol = ""
 If mvarTabla <> "" Then
   'teniendo las primarias y las normales, aca se elabora la clase
   For e = 0 To rsColumnas.Fields.Count - 1
     If rsPrimaria.RecordCount > 0 Then rsPrimaria.MoveFirst
     CampoesPK = False
     Do Until rsPrimaria.EOF = True
      ''si es primaria se le da un trato diferente
        If rsColumnas(e).Name = rsPrimaria("column_name") Then
          CampoesPK = True
          Exit Do
        End If
        rsPrimaria.MoveNext
     Loop
      If CampoesPK = True Then
        strCompararCamposLlave = strCompararCamposLlave & " mvar" & rsColumnas(e).Name & " >< """" and"
        strCamposllaveaparametrosSelect = strCamposllaveaparametrosSelect & Chr(9) & "strCamposBus = strCamposBus & """ & Chr(32) & rsColumnas(e).Name & " = ? and""" & Chr(13)
        strCamposllaveaparametrosSelect = strCamposllaveaparametrosSelect & Chr(9) & "Set par = cmd.CreateParameter(, " & IIf(rsColumnas(e).Type = adNumeric, adSingle, rsColumnas(e).Type) & ", adParamInput, " & rsColumnas(e).DefinedSize & ", mvar" & rsColumnas(e).Name & ")" & Chr(13) & Chr(9) & "cmd.Parameters.Append par" & Chr(13)
      End If
      If mvarSQLObtener = "" Then
        strCamposaparametrosSelect = strCamposaparametrosSelect & Chr(9) & "if mvar" & rsColumnas(e).Name & "><"""" then" & Chr(13)
        If rsColumnas(e).Type = adChar Or rsColumnas(e).Type = adVarChar Or rsColumnas(e).Type = adVarWChar Or rsColumnas(e).Type = adWChar Then
          strCamposaparametrosSelect = strCamposaparametrosSelect & Chr(9) & Chr(9) & "strCamposBus = strCamposBus & "" Upper(" & rsColumnas(e).Name & ") like upper( ? ) and""" & Chr(13)
        Else
          strCamposaparametrosSelect = strCamposaparametrosSelect & Chr(9) & Chr(9) & "strCamposBus = strCamposBus & """ & Chr(32) & rsColumnas(e).Name & " = ? and""" & Chr(13)
        End If
        strCamposaparametrosSelect = strCamposaparametrosSelect & Chr(9) & Chr(9) & "Set par = cmd.CreateParameter(, " & IIf(rsColumnas(e).Type = adNumeric, adSingle, rsColumnas(e).Type) & ", adParamInput, " & rsColumnas(e).DefinedSize & ", mvar" & rsColumnas(e).Name & ")" & Chr(13) & Chr(9) & "cmd.Parameters.Append par" & Chr(13) & Chr(9) & "End If" & Chr(13)
      End If
      strCamposaparametrosUpdate = strCamposaparametrosUpdate & Chr(9) & "if mvar" & rsColumnas(e).Name & "><"""" then" & Chr(13) & Chr(9) & "strCampos = strCampos & "" " & rsColumnas(e).Name & " = ?, """ & Chr(13)
      strCamposaparametrosUpdate = strCamposaparametrosUpdate & Chr(9) & Chr(9) & "Set par = cmd.CreateParameter(, " & IIf(rsColumnas(e).Type = adNumeric, adSingle, rsColumnas(e).Type) & ", adParamInput, " & rsColumnas(e).DefinedSize & ", mvar" & rsColumnas(e).Name & ")" & Chr(13) & Chr(9) & "cmd.Parameters.Append par" & Chr(13) & Chr(9) & "End If" & Chr(13)
  
      'para el insert todas son por igual
      strCamposaparametrosInsert = strCamposaparametrosInsert & Chr(9) & "if mvar" & rsColumnas(e).Name & "><"""" then" & Chr(13) & Chr(9) & "strCampos = strCampos & "" " & rsColumnas(e).Name & ",""" & Chr(13) & Chr(9) & "strValores = strValores & ""?,""" & Chr(13)
      strCamposaparametrosInsert = strCamposaparametrosInsert & Chr(9) & Chr(9) & "Set par = cmd.CreateParameter(, " & IIf(rsColumnas(e).Type = adNumeric, adSingle, rsColumnas(e).Type) & ", adParamInput, " & rsColumnas(e).DefinedSize & ", mvar" & rsColumnas(e).Name & ")" & Chr(13) & Chr(9) & "cmd.Parameters.Append par" & Chr(13) & Chr(9) & "End If" & Chr(13)
      'campos del recordset
      If e Mod 6 = 0 And e <> 0 Then
        strCamposRecordset = strCamposRecordset & " _" & Chr(13)
        strCamposParRecordset = strCamposParRecordset & " _" & Chr(13)
      End If
      If strCamposRecordset <> "" Then
        strCamposRecordset = strCamposRecordset & ","
      End If
      If strCamposParRecordset <> "" Then
        strCamposParRecordset = strCamposParRecordset & ","
      End If
      strCamposParRecordset = strCamposParRecordset & "mvar" & rsColumnas(e).Name
      'verifica si el campo puede traer valores nulos
      'si lo puede traer cambie el null por ""
      rsEstrucCol.MoveFirst
      Do Until rsEstrucCol.EOF = True
        If rsColumnas(e).Name = rsEstrucCol(3) Then
          If rsEstrucCol(10) = False Then
            strCamposRecordset = strCamposRecordset & "rsdGeneral(" & Chr(34) & rsColumnas(e).Name & Chr(34) & ")"
          Else
            strCamposRecordset = strCamposRecordset & "iif(isnull(rsdGeneral(" & Chr(34) & rsColumnas(e).Name & Chr(34) & "))," & Chr(34) & Chr(34) & ",rsdGeneral(" & Chr(34) & rsColumnas(e).Name & Chr(34) & "))"
          End If
          Exit Do
        End If
        rsEstrucCol.MoveNext
      Loop
      'esto es para usar la coleccion
      strMapeoCamposRecordSetCol = strMapeoCamposRecordSetCol & "objNewMember." & rsColumnas(e).Name & " = " & "mvar" & rsColumnas(e).Name & Chr(13) & Chr(9)
      'propiedades
      strPrivates = strPrivates & " Private mvar" & rsColumnas(e).Name & " As String 'copia local" & Chr(13)
      strPropertys = strPropertys & "Public Property Get " & rsColumnas(e).Name & "() As String" & Chr(13) & Chr(9) & rsColumnas(e).Name & " = mvar" & rsColumnas(e).Name & Chr(13) & "End Property" & Chr(13) & Chr(13)
      strPropertys = strPropertys & "Public Property Let " & rsColumnas(e).Name & "(value As String) " & Chr(13) & Chr(9) & "mvar" & rsColumnas(e).Name & " =  value" & Chr(13) & "End Property" & Chr(13) & Chr(13)
   Next
 End If
  'si tiene un select en la propiedad SQLObtener entonces use los campos del select
  'sin embargo aún puede filtrar por la llave
  If mvarSQLObtener <> "" Then
    strCamposaparametrosSelect = ""
    strCamposRecordset = ""
    strCamposParRecordset = ""
    strMapeoCamposRecordSetCol = ""
    For e = 0 To rsSelect.Fields.Count - 1
      'agrega las propiedades y condiciones adicionales para el obtener, no incluye las propiedades ya escritas por el agregar,insertar o modificar
        strCamposaparametrosSelect = strCamposaparametrosSelect & Chr(9) & "if mvar" & rsSelect(e).Name & "><"""" then" & Chr(13)
        If rsSelect(e).Type = adChar Or rsSelect(e).Type = adVarChar Or rsSelect(e).Type = adVarWChar Or rsSelect(e).Type = adWChar Then
          strCamposaparametrosSelect = strCamposaparametrosSelect & Chr(9) & Chr(9) & "strCamposBus = strCamposBus & "" Upper(" & rsSelect(e).Name & ") like upper( ? ) and""" & Chr(13)
        Else
          strCamposaparametrosSelect = strCamposaparametrosSelect & Chr(9) & Chr(9) & "strCamposBus = strCamposBus & """ & Chr(32) & rsSelect(e).Name & " = ? and""" & Chr(13)
        End If
        strCamposaparametrosSelect = strCamposaparametrosSelect & Chr(9) & Chr(9) & "Set par = cmd.CreateParameter(, " & IIf(rsSelect(e).Type = adNumeric, adSingle, rsSelect(e).Type) & ", adParamInput, " & rsSelect(e).DefinedSize & ", mvar" & rsSelect(e).Name & ")" & Chr(13) & Chr(9) & "cmd.Parameters.Append par" & Chr(13) & Chr(9) & "End If" & Chr(13)
        'campos del recordset
        If e Mod 6 = 0 And e <> 0 Then
          strCamposRecordset = strCamposRecordset & " _" & Chr(13)
          strCamposParRecordset = strCamposParRecordset & " _" & Chr(13)
        End If
        If strCamposRecordset <> "" Then
          strCamposRecordset = strCamposRecordset & ","
        End If
        If strCamposParRecordset <> "" Then
          strCamposParRecordset = strCamposParRecordset & ","
        End If
        strCamposParRecordset = strCamposParRecordset & "mvar" & rsSelect(e).Name
        'verifica si el campo puede traer valores nulos
        'si lo puede traer cambie el null por ""
        If mvarTabla <> "" Then
          rsEstrucCol.MoveFirst
          swCampoTablaPrincipal = False
          Do Until rsEstrucCol.EOF = True
            If rsSelect(e).Name = rsEstrucCol(3) Then
              swCampoTablaPrincipal = True
              If rsEstrucCol(10) = False Then
                strCamposRecordset = strCamposRecordset & "rsdGeneral(" & Chr(34) & rsSelect(e).Name & Chr(34) & ")"
              Else
                strCamposRecordset = strCamposRecordset & "iif(isnull(rsdGeneral(" & Chr(34) & rsSelect(e).Name & Chr(34) & "))," & Chr(34) & Chr(34) & ",rsdGeneral(" & Chr(34) & rsSelect(e).Name & Chr(34) & "))"
              End If
              Exit Do
            End If
            rsEstrucCol.MoveNext
          Loop
          If swCampoTablaPrincipal = False Then
            strCamposRecordset = strCamposRecordset & "iif(isnull(rsdGeneral(" & Chr(34) & rsSelect(e).Name & Chr(34) & "))," & Chr(34) & Chr(34) & ",rsdGeneral(" & Chr(34) & rsSelect(e).Name & Chr(34) & "))"
          End If
        Else
          strCamposRecordset = strCamposRecordset & "iif(isnull(rsdGeneral(" & Chr(34) & rsSelect(e).Name & Chr(34) & "))," & Chr(34) & Chr(34) & ",rsdGeneral(" & Chr(34) & rsSelect(e).Name & Chr(34) & "))"
        End If
        'esto es para usar la coleccion
        strMapeoCamposRecordSetCol = strMapeoCamposRecordSetCol & "objNewMember." & rsSelect(e).Name & " = " & "mvar" & rsSelect(e).Name & Chr(13) & Chr(9)
        propAgregarClase = True
        If mvarTabla <> "" Then
          For k = 0 To rsColumnas.Fields.Count - 1
            If rsSelect(e).Name = rsColumnas(k).Name Then
              'encontrada, no agregar la propiedad, ni la condicion
              propAgregarClase = False
              Exit For
            End If
          Next
        End If
      'si la propiedad no fue encontrada, agregela
      If propAgregarClase = True Then
        'propiedades
        strPrivates = strPrivates & " Private mvar" & rsSelect(e).Name & " As String 'copia local" & Chr(13)
        strPropertys = strPropertys & "Public Property Get " & rsSelect(e).Name & "() As String" & Chr(13) & Chr(9) & rsSelect(e).Name & " = mvar" & rsSelect(e).Name & Chr(13) & "End Property" & Chr(13) & Chr(13)
        strPropertys = strPropertys & "Public Property Let " & rsSelect(e).Name & "(value As String) " & Chr(13) & Chr(9) & "mvar" & rsSelect(e).Name & " =  value" & Chr(13) & "End Property" & Chr(13) & Chr(13)
      End If
    Next
  End If
  
  If mvarTabla <> "" Then
    strCompararCamposLlave = Left(strCompararCamposLlave, Len(strCompararCamposLlave) - 3)
  End If
  
  'mvarSQLObtener = ""
  If mvarSQLObtener = "" Then
    mvarSQLObtener = """select * from " & mvarTabla & """"
  Else
    mvarSQLObtener = """" & Replace(mvarSQLObtener, vbCrLf, Chr(34) & " _" & Chr(13) & Chr(9) & "& " & Chr(34)) & """"
  End If
  
  'coloca un nombre de objeto mejorado
  mvarTabla = IIf(mvarTabla = "", "Consulta_SQL", mvarTabla)
  
  strNombreClase = UCase(Mid(mvarTabla, 1, 1)) & LCase(Mid(mvarTabla, 2, Len(mvarTabla)))
  Do Until InStr(1, strNombreClase, "_") = 0
    strNombreClase = Mid(strNombreClase, 1, InStr(1, strNombreClase, "_") - 1) & UCase(Mid(strNombreClase, InStr(1, strNombreClase, "_") + 1, 1)) & Mid(strNombreClase, InStr(1, strNombreClase, "_") + 2, Len(strNombreClase))
  Loop
  
  strCadFunction = Replace(Replace(strPlantilla, "@Privates@", strPrivates), "@Propertys@", strPropertys)
  strCadFunction = Replace(strCadFunction, "@NombreClase@", strNombreClase)
  strCadFunction = Replace(Replace(strCadFunction, "@Tabla@", mvarTabla), "@compararcamposllave@", strCompararCamposLlave)
  strCadFunction = Replace(Replace(strCadFunction, "@camposaparametrosSelect@", strCamposaparametrosSelect), "@camposllaveaparametrosSelect@", strCamposllaveaparametrosSelect)
  strCadFunction = Replace(Replace(strCadFunction, "@camposaparametrosInsert@", strCamposaparametrosInsert), "@camposaparametrosUpdate@", strCamposaparametrosUpdate)
  strCadFunction = Replace(Replace(strCadFunction, "@SelectObtener@", mvarSQLObtener), "@camposrecordset@", strCamposRecordset)
  strCadFunction = Replace(strCadFunction, "@CadenaConn@", Replace(mvarCadenaConn, Chr(34), Chr(34) & Chr(34)))
  
  strCadcol = Replace(Replace(Replace(strPlantillaCol, "@mapeocamposrecordsetobjeto@", strMapeoCamposRecordSetCol), "@camposparrecordset@", strCamposParRecordset), "@Tabla@", mvarTabla)
  strCadcol = Replace(strCadcol, "@NombreClase@", strNombreClase)
  
  If mvarRutaArchivoCLS <> "" Then
    Open mvarRutaArchivoCLS For Output As #1
      Print #1, strCadFunction
    Close #1
    Open Left(mvarRutaArchivoCLS, Len(mvarRutaArchivoCLS) - InStr(1, StrReverse(mvarRutaArchivoCLS), "\")) & "\col" & Right(mvarRutaArchivoCLS, InStr(1, StrReverse(mvarRutaArchivoCLS), "\") - 1) For Output As #1
      Print #1, strCadcol
    Close #1
    
  End If
  
  GenerarClase = strCadFunction
    
  
'  Open App.Path + "\pruebas.cls" For Output As #1
'    Print #1, strCadFunction
'  Close #1
  If Not (rsColumnas Is Nothing) Then
    If rsColumnas.State <> 0 Then rsColumnas.Close
  End If
  If Not (rsPrimaria Is Nothing) Then
    If rsPrimaria.State <> 0 Then rsPrimaria.Close
  End If
  Set rsColumnas = Nothing
  Set rsPrimaria = Nothing
  Set conn = Nothing
Exit Function
er:
  If Not (rsColumnas Is Nothing) Then
    If rsColumnas.State <> 0 Then rsColumnas.Close
  End If
  If Not (rsPrimaria Is Nothing) Then
    If rsPrimaria.State <> 0 Then rsPrimaria.Close
  End If
  Set rsColumnas = Nothing
  Set rsPrimaria = Nothing
If conn.State <> 0 Then conn.Close
Set conn = Nothing
Set rs = Nothing
mvarMensajeError = Err.Description
Exit Function
Resume
End Function

Public Property Get MensajeError() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.MensajeError
    MensajeError = mvarMensajeError
End Property



Public Property Let CadenaConn(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CadenaConn = 5
    mvarCadenaConn = vData
End Property


Public Property Get CadenaConn() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CadenaConn
    CadenaConn = mvarCadenaConn
End Property



Public Property Let Schema(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Schema = 5
    mvarSchema = vData
End Property


Public Property Get Schema() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Schema
    Schema = mvarSchema
End Property



Public Property Let SQLObtener(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.SQLObtener = 5
    mvarSQLObtener = vData
End Property


Public Property Get SQLObtener() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.SQLObtener
    SQLObtener = mvarSQLObtener
End Property



Public Property Let Tabla(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Tabla = 5
    mvarTabla = vData
End Property


Public Property Get Tabla() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Tabla
    Tabla = mvarTabla
End Property



