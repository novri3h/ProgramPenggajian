Attribute VB_Name = "MdlGaji"

Public Conn As New ADODB.Connection
Public RSPerkiraan As ADODB.Recordset
Public RSDetail As ADODB.Recordset
Public RSGaji As ADODB.Recordset
Public RSKasir As ADODB.Recordset
Public RSPegawai As ADODB.Recordset

Public Sub BukaDB()
Set Conn = New ADODB.Connection
Set RSPerkiraan = New ADODB.Recordset
Set RSDetail = New ADODB.Recordset
Set RSGaji = New ADODB.Recordset
Set RSKasir = New ADODB.Recordset
Set RSPegawai = New ADODB.Recordset
Conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ADOGaji.mdb"
End Sub

