VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales para almacenar los valores de las propiedades
Private mvarCadenaConn As String 'copia local
Private mvarNombreObjeto As String 'copia local
Private mvarMensajeError As String 'copia local
'Para activar este evento, use RaiseEvent con la siguiente sintaxis:
'RaiseEvent ObteniendoCadenaConn[(arg1, arg2, ... , argn)]
Public Event ObteniendoCadenaConn()
Public Event ObteniendoListadoTablas(porcentaje As Integer)
Public Event ObteniendoListadoSchemas(porcentaje As Integer)
Private mvarTipoObjeto As String 'copia local
Private mvarObtenerObtieneRegistros As Boolean 'copia local

Public Function ObtenerListadoSchemas() As colTablas
Dim conn As ADODB.Connection, rst As ADODB.Recordset
Dim por As Integer
Dim coltab As colTablas
On Local Error GoTo er:
  'requisitos
  mvarMensajeError = ""
  If mvarCadenaConn = "" Then
    mvarMensajeError = "Debe diligenciar la cadena de conexión antes de obtener el listado de esquemas"
  Else
    Set conn = New ADODB.Connection
    conn.CursorLocation = adUseClient
    conn.Open mvarCadenaConn
    ' Recupera las tablas de la base de datos mediante OpenSchema
    Set rst = conn.OpenSchema(adSchemaSchemata)
    'desconecta el recordset
    Set rst.ActiveConnection = Nothing
    conn.Close
    Set coltab = New colTablas
    Do Until rst.EOF
      coltab.Add mvarCadenaConn, rst("SCHEMA_NAME"), "ESQUEMA"
      por = (rst.AbsolutePosition / rst.RecordCount) * 100
      RaiseEvent ObteniendoListadoSchemas(por)
      rst.MoveNext
    Loop
  End If
  Set ObtenerListadoSchemas = coltab
  If rst.State <> 0 Then rst.Close
  Set conn = Nothing
  Set rst = Nothing
  Set coltab = Nothing
  Exit Function
er:
mvarMensajeError = Err.Description
End Function

Public Function ObtenerListadoTablas() As colTablas
Dim conn As ADODB.Connection, rst As ADODB.Recordset
Dim por As Integer
Dim coltab As colTablas
On Local Error GoTo er:
  'requisitos
  mvarMensajeError = ""
  If mvarCadenaConn = "" Then
    mvarMensajeError = "Debe diligenciar la cadena de conexión antes de obtener el listado de tablas"
  'ElseIf mvarNombreObjeto = "" Then
  '  mvarMensajeError = "Debe diligenciar el nombre del schema en la propiedad 'NombreObjeto' antes de obtener el listado de tablas"
  Else
    Set conn = New ADODB.Connection
    conn.CursorLocation = adUseClient
    conn.Open mvarCadenaConn
    ' Recupera las tablas de la base de datos mediante OpenSchema
    Set rst = conn.OpenSchema(adSchemaTables, Array(Empty, IIf(mvarNombreObjeto = "", Empty, mvarNombreObjeto), Empty, "TABLE"))
    'desconecta el recordset
    Set rst.ActiveConnection = Nothing
    conn.Close
    Set coltab = New colTablas
    Do Until rst.EOF
      coltab.Add mvarCadenaConn, rst("Table_Name"), "TABLA"
      por = (rst.AbsolutePosition / rst.RecordCount) * 100
      RaiseEvent ObteniendoListadoTablas(por)
      rst.MoveNext
    Loop
    If rst.State <> 0 Then rst.Close
  End If
  Set ObtenerListadoTablas = coltab
  Set conn = Nothing
  Set rst = Nothing
  Set coltab = Nothing
  Exit Function
er:
mvarMensajeError = Err.Description
End Function

Public Property Get ObtenerObtieneRegistros() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ObtenerObtieneRegistros
    ObtenerObtieneRegistros = mvarObtenerObtieneRegistros
End Property

Public Property Get MensajeError() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.MensajeError
    MensajeError = mvarMensajeError
End Property


Public Property Let tipoObjeto(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.NombreObjeto = 5
    mvarTipoObjeto = vData
End Property


Public Property Get tipoObjeto() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.NombreObjeto
    tipoObjeto = mvarTipoObjeto
End Property


Public Property Let NombreObjeto(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.NombreObjeto = 5
    mvarNombreObjeto = vData
End Property


Public Property Get NombreObjeto() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.NombreObjeto
    NombreObjeto = mvarNombreObjeto
End Property

Public Function ObtenerCadenaConn(Optional Verificar As Boolean = True) As String
On Local Error GoTo er:
Dim o_DataLinks As DataLinks
Dim conn As ADODB.Connection
Dim objpropcadconn As Object
mvarMensajeError = ""
Set o_DataLinks = New DataLinks
RaiseEvent ObteniendoCadenaConn
'error controlado
On Local Error Resume Next
Set conn = New ADODB.Connection
If mvarCadenaConn = "" Then
  ObtenerCadenaConn = o_DataLinks.PromptNew
Else
  'ObtenerCadenaConn = o_DataLinks.PromptEdit(mvarCadenaConn)
  If Err.Description <> "" Then
    ObtenerCadenaConn = o_DataLinks.PromptNew
  Else
    conn.ConnectionString = mvarCadenaConn
    If o_DataLinks.PromptEdit(conn) Then
      ObtenerCadenaConn = conn.ConnectionString
    End If
  End If
End If
On Local Error GoTo er:
If Verificar = True Then
  conn.Open ObtenerCadenaConn
  conn.Close
End If
Set conn = Nothing
Set o_DataLinks = Nothing
Exit Function
er:
ObtenerCadenaConn = ""
mvarMensajeError = Err.Description
End Function

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



