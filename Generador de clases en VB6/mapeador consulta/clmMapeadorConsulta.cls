VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmMapeadorConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'variables locales para almacenar los valores de las propiedades
Private mvarCadenaConn As String 'copia local
Private mvarParametros As Collection 'copia local
Private mvarMensajeError As String 'copia local

Public Property Get MensajeError() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Parametros
    Set MensajeError = mvarMensajeError
End Property

Public Function EjecutarConsulta(Consulta As String) As Collection
  On Local Error GoTo er:
  'acceso a datos
  Dim conn As ADODB.Connection
  Dim rs As ADODB.Recordset
  Dim cmd As ADODB.Command
  
  'resultados
  Dim reg As Collection
  Dim camp As Collection
  
  Set reg = New Collection
  Set conn = New Connection
  Set cmd = New ADODB.Command
  conn.Open mvarCadenaConn
  cmd.ActiveConnection = conn
  
  For Each par In mvarParametros
    cmd.Parameters.Append cmd.CreateParameter(, adVarChar, adParamInput, 4000, par)
  Next
  
  cmd.CommandText = Consulta
  Set rs = cmd.Execute
  
  Do Until rs.EOF = True
    Set camp = New Collection
    For e = 0 To rs.Fields.Count - 1
      camp.Add rs(e)
    Next
    reg.Add camp
    Set camp = Nothing
    rs.MoveNext
  Loop
  
  If conn.State <> 0 Then conn.Close
  If rs.State <> 0 Then rs.Close
  Set rs = Nothing
  Set conn = Nothing
  Set EjecutarConsulta = reg
  Set reg = Nothing
  Exit Function
er:
  If conn.State <> 0 Then conn.Close
  If rs.State <> 0 Then rs.Close
  Set rs = Nothing
  
  Set conn = Nothing
  Set camp = Nothing
  Set cmd = Nothing
  mvarMensajeError = Err.Description
End Function

Public Property Set Parametros(ByVal vData As Collection)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Parametros = Form1
    Set mvarParametros = vData
End Property


Public Property Get Parametros() As Collection
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Parametros
    Set Parametros = mvarParametros
End Property

Public Sub LimpiarParametros()
  Do Until mvarParametros.Count = 0
    mvarParametros.Remove 1
  Loop
End Sub

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

Private Sub Class_Initialize()
  Set mvarParametros = New Collection
End Sub
