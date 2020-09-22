VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmGeneradorClases 
   Caption         =   "Generador de clases para tablas de BD en VB6"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog c1 
      Left            =   4920
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame fraRutaCLS 
      Caption         =   "5. Seleccione la ruta donde será guardado el archivo (opcional)"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   7095
      Begin VB.TextBox txtRutaCLS 
         Height          =   405
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   6255
      End
      Begin VB.CommandButton cmdSelGuardarCLS 
         Caption         =   "..."
         Height          =   375
         Left            =   6600
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame fraschema 
      Caption         =   "2. Seleccione el esquema de la base de datos seleccionada"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   7095
      Begin VB.ListBox lstEsquemaBD 
         Height          =   645
         ItemData        =   "frmGeneradorClases.frx":0000
         Left            =   240
         List            =   "frmGeneradorClases.frx":0002
         TabIndex        =   9
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdGenerarClase 
      Caption         =   "Generar clase de VB6 con metodos éstandar"
      Height          =   975
      Left            =   2760
      Picture         =   "frmGeneradorClases.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Frame fraSelTabVisBD 
      Caption         =   "3. Seleccione la tabla de la BD"
      Enabled         =   0   'False
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   7095
      Begin VB.ListBox lstTabla 
         Height          =   1620
         ItemData        =   "frmGeneradorClases.frx":0586
         Left            =   240
         List            =   "frmGeneradorClases.frx":0588
         TabIndex        =   4
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Frame fraSelDS 
      Caption         =   "1. Seleccione el origen de datos"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdGenCadConn 
         Caption         =   "..."
         Height          =   375
         Left            =   6600
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCadConn 
         Height          =   405
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Frame fraSQLObtener 
      Caption         =   "4. Escriba el SQL para el metodo de Obtener datos de la tabla seleccionada (opcional)"
      Enabled         =   0   'False
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   7095
      Begin RichTextLib.RichTextBox txtSQLObtener 
         Height          =   1575
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2778
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmGeneradorClases.frx":058A
      End
      Begin VB.ComboBox cmbSugerencias 
         Height          =   315
         Left            =   4680
         TabIndex        =   13
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdGenSQLBuilder 
         Caption         =   "&Generar SQL por diseñador"
         Height          =   855
         Left            =   5520
         Picture         =   "frmGeneradorClases.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Menu mnuOpcionesEspeciales 
      Caption         =   "OpcionesEspeciales"
      Visible         =   0   'False
      Begin VB.Menu mnuOEMenuSQL 
         Caption         =   "MenuSQL"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "frmGeneradorClases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pSQLParser As MSSQLParser.vbSQLParser
Dim tablas As clsTablas
Dim swNorevisar As Boolean

Private Sub Form_Load()
    Set pSQLParser = New MSSQLParser.vbSQLParser
End Sub

Private Sub cmbsugerencias_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Or KeyAscii = 13 Then
  cadini = Mid(txtSQLObtener.Text, 1, txtSQLObtener.SelStart)
  cadfin = Mid(txtSQLObtener.Text, txtSQLObtener.SelStart + 1, Len(txtSQLObtener.Text))
  txtSQLObtener.Text = cadini & cmbSugerencias & cadfin
  cmbSugerencias.Visible = False
  txtSQLObtener.SetFocus
  txtSQLObtener.SelStart = Len(cadini & cmbSugerencias)
ElseIf KeyAscii = vbKeyEscape Then
  cmbSugerencias.Visible = False
  txtSQLObtener.SetFocus
  txtSQLObtener.SelStart = Len(cadini & cmbSugerencias)
End If
End Sub

Private Sub cmdGenCadConn_Click()
ObtenerSchemas
If lstEsquemaBD.ListCount > 0 Then lstEsquemaBD.ListIndex = 0
If ObtenerTablas = False Then
  fraSelTabVisBD.Enabled = False
  fraSQLObtener.Enabled = False
Else
  fraSelTabVisBD.Enabled = True
  fraSQLObtener.Enabled = True
End If
End Sub

Private Sub cmdGenerarClase_Click()
Dim generadorClase As clsGeneradorClase
Set generadorClase = New clsGeneradorClase
  generadorClase.CadenaConn = txtCadConn
  generadorClase.Schema = lstEsquemaBD.Text
  generadorClase.Tabla = lstTabla
  generadorClase.SQLObtener = Trim(txtSQLObtener.Text)
  generadorClase.RutaArchivoCLS = txtRutaCLS
  frmTextoClase.txtTextoClaseGen = generadorClase.GenerarClase
  If generadorClase.MensajeError <> "" Then
    MsgBox "Error al tratar de generar la clase. El error es: '" & generadorClase.MensajeError & "'", vbExclamation
  ElseIf generadorClase.RutaArchivoCLS <> "" Then
    MsgBox "Clase generada con éxito en el archivo " & txtRutaCLS, vbInformation
  Else
    MsgBox "Clase generada con éxito. Acá esta el código", vbInformation
    frmTextoClase.Show 1
  End If
Set generadorClase = Nothing
End Sub

Private Sub cmdGenSQLBuilder_Click()
'no implementado por ser lento
'  Dim fqb As frmQuerybuilder
'  Set fqb = New frmQuerybuilder
'  fqb.ActiveQueryBuilderX1.ConnectionString = txtCadConn
'  fqb.ActiveQueryBuilderX1.SQL = txtSQLObtener
'  frmQuerybuilder.Show 1
End Sub

Function ObtenerSchemas() As Boolean
'para los resultados de la coleccion
Dim coltabla As colTablas
Dim tb As clsTablas
'Obtiene la cadena de conexión
Set tablas = New clsTablas
tablas.CadenaConn = txtCadConn
txtCadConn = tablas.ObtenerCadenaConn(True)
If tablas.MensajeError <> "" Then
  MsgBox "Ocurrio un error al tratar de obtener la cadena de conexión. El error fue: " & tablas.MensajeError
  ObtenerSchemas = False
Else
  MsgBox "Conexión efectuada con éxito", vbInformation, Me.Caption
  tablas.CadenaConn = txtCadConn
  Set coltabla = tablas.ObtenerListadoSchemas
  If tablas.MensajeError <> "" Then
    MsgBox "Ocurrio un error al tratar de obtener la lista de esquemas. Se intentará obtener el listado de tablas sin esquema. El error fue: " & tablas.MensajeError, vbExclamation
    ObtenerSchemas = False
  Else
    lstEsquemaBD.Clear
    For Each tb In coltabla
      lstEsquemaBD.AddItem tb.NombreObjeto
    Next
    ObtenerSchemas = True
  End If
End If
Set tablas = Nothing
Set tb = Nothing
Set coltabla = Nothing
End Function

Function ObtenerTablas() As Boolean
'para los resultados de la coleccion
Dim coltabla As colTablas
Dim tb As clsTablas
'Obtiene la cadena de conexión
Set tablas = New clsTablas
'tablas.CadenaConn = txtCadConn
'txtCadConn = tablas.ObtenerCadenaConn(True)
'If tablas.MensajeError <> "" Then
'  MsgBox "Ocurrio un error al tratar de obtener la cadena de conexión. El error fue: " & tablas.MensajeError
'  ObtenerTablas = False
'Else
  tablas.CadenaConn = txtCadConn
  tablas.NombreObjeto = lstEsquemaBD.Text
  Set coltabla = tablas.ObtenerListadoTablas
  If tablas.MensajeError <> "" Then
    MsgBox "Ocurrio un error al tratar de obtener la lista de tablas. El error fue: " & tablas.MensajeError
    ObtenerTablas = False
  Else
    lstTabla.Clear
    For Each tb In coltabla
      lstTabla.AddItem tb.NombreObjeto
    Next
    ObtenerTablas = True
  End If
'End If
Set tb = Nothing
Set tablas = Nothing
Set coltabla = Nothing
End Function

Private Sub cmdSelGuardarCLS_Click()
On Local Error Resume Next
  c1.Filter = "Modulo de clase | *.cls"
  c1.ShowOpen
  txtRutaCLS = c1.FileName
End Sub

Private Sub lstEsquemaBD_Click()
If ObtenerTablas = False Then
  fraSelTabVisBD.Enabled = False
  fraSQLObtener.Enabled = False
Else
  fraSelTabVisBD.Enabled = True
  fraSQLObtener.Enabled = True
End If
End Sub

Private Sub lstTabla_DblClick()
cadini = Mid(txtSQLObtener.Text, 1, txtSQLObtener.SelStart)
cadfin = Mid(txtSQLObtener.Text, txtSQLObtener.SelStart + 1, Len(txtSQLObtener.Text))
txtSQLObtener.Text = cadini & lstTabla.Text & cadfin
txtSQLObtener.SetFocus
txtSQLObtener.SelStart = Len(cadini & lstTabla.Text)
End Sub

Private Sub lstTabla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lstTabla_DblClick
End Sub

Private Sub lstTabla_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then lstTabla.ListIndex = -1
End Sub

Private Sub mnuOEMenuSQL_Click()
'saca las tablas en el combo
cmbSugerencias.Clear
For e = 0 To cmbSugerencias.ListCount - 1
  cmbSugerencias.AddItem lstTabla(e)
Next
px = ModPosicionCursorTextbox.GetCurrentCol(txtSQLObtener)
py = ModPosicionCursorTextbox.GetCurrentLine(txtSQLObtener)
cmbSugerencias.Move ((px * 105) - 450 + txtSQLObtener.Left), (py * txtSQLObtener.Font.Size * 15) + txtSQLObtener.Top
cmbSugerencias.ZOrder 0
cmbSugerencias.Visible = True
cmbSugerencias.ListIndex = 0
cmbSugerencias.SetFocus

End Sub

'para colorear palabras clave
Private Sub txtSQLObtener_Change()
    '/cf1   Select, Insert
    '/cf2   Data type
    '/cf3   Functions (Isnull, RowNum)
    '/cf4
    '/cf5   Logic Operators (And, Or)
    '/cf6   Text between ''
    '/cf7   Numeric values
    '/cf8   Comments
    Dim temp As String
    Dim X As Integer
    If swNorevisar = False Then
      X = txtSQLObtener.SelStart
      temp = pSQLParser.ParseSQLSyntax(txtSQLObtener.Text, vbOracleSyntax)
          temp = "{\colortbl;" _
          & "\red0\green0\blue255;" _
          & "\red255\green0\blue0;" _
          & "\red255\green0\blue255;" _
          & "\red0\green255\blue0;" _
          & "\red128\green128\blue128;" _
          & "\red255\green0\blue0;" _
          & "\red153\green0\blue204;" _
          & "\red0\green150\blue100;" _
          & "\red0\green255\blue0;" _
          & "\red0\green0\blue0;}" _
          & temp
      temp = "{" & temp & "}"
      temp = Replace$(temp, "\cf10", "{##0")
      temp = Replace$(temp, "\cf1", "{##1")
      temp = Replace$(temp, "\cf2", "{##2")
      temp = Replace$(temp, "\cf3", "{##3")
      temp = Replace$(temp, "\cf4", "{##4")
      temp = Replace$(temp, "\cf5", "{##5")
      temp = Replace$(temp, "\cf6", "{##6")
      temp = Replace$(temp, "\cf7", "{##7")
      temp = Replace$(temp, "\cf8", "{##8")
      temp = Replace$(temp, "\cf9", "{##9")
      temp = Replace$(temp, "\cf ", "}", , , vbBinaryCompare)
      temp = Replace$(temp, "##1", "\cf1", , , vbBinaryCompare)
      temp = Replace$(temp, "##2", "\cf2", , , vbBinaryCompare)
      temp = Replace$(temp, "##3", "\cf3", , , vbBinaryCompare)
      temp = Replace$(temp, "##4", "\cf4", , , vbBinaryCompare)
      temp = Replace$(temp, "##5", "\cf5", , , vbBinaryCompare)
      temp = Replace$(temp, "##6", "\cf6", , , vbBinaryCompare)
      temp = Replace$(temp, "##7", "\cf7", , , vbBinaryCompare)
      temp = Replace$(temp, "##8", "\cf8", , , vbBinaryCompare)
      temp = Replace$(temp, "##9", "\cf9", , , vbBinaryCompare)
      temp = Replace$(temp, "##0", "\cf10", , , vbBinaryCompare)
      'temp = Replace$(temp, "\par", "\par \par", , , vbBinaryCompare)
      
      txtSQLObtener.TextRTF = temp
      txtSQLObtener.SelStart = X
    End If
End Sub

Private Sub txtSQLObtener_KeyPress(KeyAscii As Integer)
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection
'sw para colorear palabras clave
swNorevisar = (KeyAscii = 13)
'lista los campos de una tabla
If Chr(KeyAscii) = "." Then
  cadini = Mid(txtSQLObtener.Text, 1, txtSQLObtener.SelStart)
  cadfin = Mid(txtSQLObtener.Text, txtSQLObtener.SelStart + 1, Len(txtSQLObtener.Text))
  strcad = StrReverse(cadini)
  poscoma = InStr(1, strcad, ",")
  posespacio = InStr(1, strcad, " ")
  If poscoma < posespacio And poscoma <> 0 Then
    posini = poscoma
  ElseIf posespacio <> 0 Then
    posini = posespacio
  End If
  If posini <> 0 Then
    strtabla = Right(cadini, posini - 1)
    Set conn = New ADODB.Connection
    conn.Open txtCadConn
    Err.Clear
    On Local Error Resume Next
    Set rs = conn.Execute("select * from " & strtabla & " where 1=2")
    If Err.Description = "" Then
      cmbSugerencias.Clear
      cmbSugerencias.AddItem "*"  'todos los campos de la tabla
      For e = 0 To rs.Fields.Count - 1
        cmbSugerencias.AddItem rs.Fields(e).Name
      Next
      px = ModPosicionCursorTextbox.GetCurrentCol(txtSQLObtener)
      py = ModPosicionCursorTextbox.GetCurrentLine(txtSQLObtener)
      cmbSugerencias.Move ((px * 105) - 450 + txtSQLObtener.Left), (py * txtSQLObtener.Font.Size * 15) + txtSQLObtener.Top
      cmbSugerencias.ZOrder 0
      cmbSugerencias.Visible = True
      cmbSugerencias.ListIndex = 0
      cmbSugerencias.SetFocus
      rs.Close
      Set rs = Nothing
    End If
  End If
End If
End Sub
