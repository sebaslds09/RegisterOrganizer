Attribute VB_Name = "Module1"
Public conexion As ADODB.Connection
Public Usuario As New ADODB.Recordset
Public cliente As New ADODB.Recordset
Public maquina As New ADODB.Recordset
Public operario As New ADODB.Recordset
Public produccion As New ADODB.Recordset
Public reporte As New ADODB.Recordset
Public hilo As New ADODB.Recordset
Public produccion_hilo As New ADODB.Recordset
Public materia_prima As New ADODB.Recordset
Public reporte_detalle As New ADODB.Recordset
Public tipo_maquina As New ADODB.Recordset
Public causa As New ADODB.Recordset
Public tiempo_muerto As New ADODB.Recordset
Public suiche As New ADODB.Recordset
Public consulta As New ADODB.Recordset
Public consulta2 As New ADODB.Recordset
Public st As Integer
Public sh As Integer
Public Function conectar() As Boolean
On Error GoTo error
 conectar = False
 Set conexion = New ADODB.Connection
 conexion.Provider = "microsoft.jet.oledb.4.0; Jet OLEDB:Database Password=92120950322sebastian"
 conexion.ConnectionString = App.Path & "\Base de datos\bordados marion.mdb"
 conexion.Open
 conectar = True
Exit Function
error:
Dim varerror As ADODB.error
    For Each varerror In conexion.Errors
        MsgBox varerror.Number & "" & varerror.Description
    Next
End Function
Sub abrirtablausuario()
If Usuario.State = 1 Then
Usuario.Close
End If
Usuario.LockType = adLockOptimistic
Usuario.Open "usuario", conexion
End Sub
Sub abrirtablacliente()
If cliente.State = 1 Then
cliente.Close
End If
cliente.LockType = adLockOptimistic
cliente.Open "cliente", conexion
End Sub
Sub abrirtablamaquina()
If maquina.State = 1 Then
maquina.Close
End If
maquina.LockType = adLockOptimistic
maquina.Open "maquina", conexion
End Sub
Sub abrirtablaoperario()
If operario.State = 1 Then
operario.Close
End If
operario.LockType = adLockOptimistic
operario.Open "operario", conexion
End Sub
Sub abrirtablaproduccion()
If produccion.State = 1 Then
produccion.Close
End If
produccion.LockType = adLockOptimistic
produccion.Open "produccion", conexion
End Sub
Sub abrirtablareporte()
If reporte.State = 1 Then
reporte.Close
End If
reporte.LockType = adLockOptimistic
reporte.Open "reporte", conexion
End Sub
Sub abrirtablareporte_detalle()
If reporte_detalle.State = 1 Then
reporte_detalle.Close
End If
reporte_detalle.LockType = adLockOptimistic
reporte_detalle.Open "reporte_detalle", conexion
End Sub
Sub abrirtablatipo_maquina()
If tipo_maquina.State = 1 Then
tipo_maquina.Close
End If
tipo_maquina.LockType = adLockOptimistic
tipo_maquina.Open "tipo_maquina", conexion
End Sub
Sub abrirtablahilo()
If hilo.State = 1 Then
hilo.Close
End If
hilo.LockType = adLockOptimistic
hilo.Open "hilo", conexion
End Sub
Sub abrirtablaproduccion_hilo()
If produccion_hilo.State = 1 Then
produccion_hilo.Close
End If
produccion_hilo.LockType = adLockOptimistic
produccion_hilo.Open "produccion_hilo", conexion
End Sub
Sub abrirtablamateria_prima()
If materia_prima.State = 1 Then
materia_prima.Close
End If
materia_prima.LockType = adLockOptimistic
materia_prima.Open "materia_prima", conexion
End Sub
Sub abrirtablacausa()
If causa.State = 1 Then
causa.Close
End If
causa.LockType = adLockOptimistic
causa.Open "causa", conexion
End Sub
Sub abrirtablatiempo_muerto()
If tiempo_muerto.State = 1 Then
tiempo_muerto.Close
End If
tiempo_muerto.LockType = adLockOptimistic
tiempo_muerto.Open "tiempo_muerto", conexion
End Sub
Sub abrirtablasuiche()
If suiche.State = 1 Then
suiche.Close
End If
suiche.LockType = adLockOptimistic
suiche.Open "suiche", conexion
End Sub
Sub cerrarconsulta()
If consulta.State = 1 Then
 consulta.Close
End If
End Sub
Sub cerrarconsulta2()
If consulta2.State = 1 Then
 consulta2.Close
End If
End Sub

