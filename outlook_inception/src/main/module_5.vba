Public Sub DbInsert(inputID As String, inputAgent As String)

Set con = CreateObject("ADODB.Connection")

Dim dbPath As String
dbPath = "Q:\23.2 Danmark\05 Mæglerservice\1. linje Team\Udvikling\app_outlook\first_line.accdb"

With con
 .Provider = "Microsoft.ACE.OLEDB.12.0"
 .Open dbPath
End With

Dim query As String

query = " INSERT INTO workdist " & "(Entry_id, agent) VALUES " & "('" & inputID & "', '" & inputAgent & "');"

con.Execute query


End Sub
