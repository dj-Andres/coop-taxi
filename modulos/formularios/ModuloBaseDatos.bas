Attribute VB_Name = "ModuloBaseDatos"
Public IDSOCIO1 As Integer


Public IDUSUARIO1 As Integer


Public CONEXION As ADODB.Connection

Public COMD  As ADODB.Command

Public Rs As ADODB.Recordset
Public Function conectardb()
    Set CONEXION = New ADODB.Connection
    Set COMD = New ADODB.Command
    Set Rs = New ADODB.Recordset
    CONEXION.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PostgreSQL37"
Set COMD.ActiveConnection = CONEXION

CONEXION.CursorLocation = adUseClient

End Function

