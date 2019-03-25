Sub assignMe()

'Getting the username of the active user
Dim userpath As String
Dim activeuser As String
userpath = CreateObject("WScript.Shell").specialfolders("Desktop")
userpath = Left(userpath, Len(userpath) - 8)
activeuser = Right(userpath, 6)

'Getting the entry_id of the selected mailitem
Dim aCellR As Integer
Dim mailItemId As String
aCellR = ActiveCell.Row
mailItemId = ActiveSheet.Range("B" & aCellR).value

'Do the INSERT in Accessdb
Call DbInsert(mailItemId, activeuser)

'Update data connections
ActiveWorkbook.RefreshAll


End Sub
