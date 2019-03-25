Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
'Update 20140318
Static xRow
Static xColumn
If xColumn <> "" Then
    With Rows(xRow).Interior
        .ColorIndex = xlNone
    End With
End If
pRow = Selection.Row
pColumn = Selection.Column
xRow = pRow
xColumn = pColumn
With Range("B" & pRow & ":J" & pRow).Interior
    .ColorIndex = 36
    .Pattern = xlSolid
End With

On Error Resume Next

Set App = CreateObject("Outlook.Application")
Set NS = App.GetNamespace("MAPI")
NS.Logon

Set Msg = NS.GetItemFromID(Range("B" & pRow).value)

ActiveSheet.Preview.Text = Msg.Body

End Sub
