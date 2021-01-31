Attribute VB_Name = "ModuloFunciones"
Public tblbuscarID As New ADODB.Recordset
Public Function letras(ByVal codigo As Integer)
    Select Case codigo
        Case 65 To 90:
        Case 97 To 122:
        Case 209:
        Case 241:
        Case 8:
        Case 32:
    Case Else
        codigo = 0
    End Select
    letras = codigo
End Function
Public Function Direccion(ByVal codigo As Integer)
    Select Case codigo
        Case 65 To 90:
        Case 48 To 57:
        Case 97 To 122:
        Case 40 To 41:
        Case 44 To 45:
        Case 209:
        Case 241:
        Case 46:
        Case 35:
        Case 47:
        Case 8:
        Case 32:
    Case Else
        codigo = 0
     End Select
     Direccion = codigo
End Function
Public Function Numeros(ByVal codigo As Integer)
    Select Case codigo
        Case 48 To 57:
        Case 45:
        Case 8:
        Case Else
            codigo = 0
        End Select
        Numeros = codigo
End Function
Public Function celular(ByVal codigo As Integer)
    Select Case codigo
        Case 48 To 57:
        Case 8:
        Case Else
            codigo = 0
        End Select
        celular = codigo
End Function
Public Function numeros_letras(ByVal codigo As Integer)
    Select Case codigo
        Case 65 To 90:
        Case 97 To 122:
        Case 48 To 57:
        Case 209:
        Case 241:
        Case 8:
        Case 32:
        Case Else
            codigo = 0
        End Select
        numeros_letras = codigo
        
End Function
Public Function buscarID(ByVal tabla As String, ByVal campo As String, ByVal valor As String) As Integer
    Dim respuestaID As Integer
    Set tblbuscarID = Nothing
    tblbuscarID.Open "select * from " & tabla & " where " & campo & " = '" & valor & "'", CONEXION, adOpenDynamic, adLockOptimistic
    If Not (tblbuscarID.EOF) Then
        tblbuscarID.MoveFirst
        respuestaID = tblbuscarID.Fields(0).Value
        buscarID = respuestaID
    End If
End Function
Public Function buscarCampo(ByVal tabla As String, ByVal campoID As String, ByVal ID As Integer, ByVal valorCampo As String) As String
    Dim respuestaID As String
    Set tblbuscarID = Nothing
    tblbuscarID.Open "select * from " & tabla & " where " & campoID & " = '" & ID & "'", CONEXION, adOpenDynamic, adLockOptimistic
    If Not (tblbuscarID.EOF) Then
        tblbuscarID.MoveFirst
        respuestaID = tblbuscarID.Fields(valorCampo).Value
        buscarCampo = respuestaID
        Set tblbuscarID = Nothing
    End If
End Function

