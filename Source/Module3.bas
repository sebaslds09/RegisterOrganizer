Attribute VB_Name = "Module3"
Sub cargarcodigo(tabla As String, frmm As Form)
cerrarconsulta
sentencia = "select max(codigo) as codigo from" & " " & tabla
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic
With frmm
If Not consulta.RecordCount = 0 Then
If Not consulta.Fields("codigo") = 0 Then
.txtcodigo.Text = consulta.Fields("codigo") + 1
Else
.txtcodigo.Text = 1
End If
End If
End With
End Sub
Sub cargaricono(Form As Form)
Form.Icon = frmmenu.Icon
End Sub
Sub consultar(sql As String)
cerrarconsulta
sentencia = sql
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic
End Sub
Sub consultarotro(sqlo As String)
cerrarconsulta
sentencia = sqlo
conexion.Execute sqlo
End Sub
Sub cargarfle(sqlf As String, title As String, campo As String, ncols As Integer, frmf As Form)
Dim n As Integer
consultar sqlf
If Not consulta.RecordCount = 0 Then
With frmf
.flenombre.AutoSetup consulta.RecordCount, ncols, True, True, title

End With
n = 1
Do Until consulta.EOF
With frmf.flenombre
.Row = n
.Col = 0
.Text = n

.Col = 1
.Text = consulta.Fields(campo)

consulta.MoveNext
n = n + 1
End With
Loop
 Else
 With frmf.flenombre
.Row = 1
.Col = 0
.Text = 0

.Col = 1
.Text = ""
End With
End If
End Sub
Sub consultar2(sql As String)
cerrarconsulta2
sentencia = sql
consulta2.Source = sentencia
consulta2.Open , conexion, adOpenStatic
End Sub

