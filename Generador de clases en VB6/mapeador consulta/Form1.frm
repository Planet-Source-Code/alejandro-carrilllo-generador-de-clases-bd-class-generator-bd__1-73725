VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim map As clmMapeadorConsulta
Set map = New clmMapeadorConsulta
Dim reg As Collection
map.LimpiarParametros
map.CadenaConn = "Provider=MSDASQL.1;Password=davivienda_des;Persist Security Info=True;User ID=davivienda_des;Extended Properties=""DRIVER={Oracle en OraClient10g_home1};SERVER=SRORACLE;UID=davivienda_des;PWD=davivienda_des;DBQ=SRORACLE;DBA=W;APA=T;EXC=F;XSM=Default;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;GDE=F;FRL=Lo;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=Me;CSR=F;FWC=F;FBS=60000;TLO=O;"""
map.Parametros.Add "%M%"
Set reg = map.EjecutarConsulta("select * from centrosdecosto where CENNOMBRE like ?")
'muestra
For Each campos In reg
  For Each c In campos
    MsgBox c
  Next
Next

End Sub
