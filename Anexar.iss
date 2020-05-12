Sub Main
	Call AppendDatabase()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Archivo: Anexar bases de datos
Function AppendDatabase
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "Ejemplox-Detalle de ventas.IMD"
	task.Criteria = " TOTAL>= 5000"
	dbName = "Anexar_01.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function
