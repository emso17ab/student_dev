Sub Knap4_Klik()

Dim aCellR, aCellRegionR, currentR As Integer
Set tbl_1 = ActiveSheet.ListObjects("Tabel1")

aCellR = ActiveCell.Row
aCellRegionR = ActiveCell.CurrentRegion.Row
currentR = aCellR - aCellRegionR
   
Dim answer As Integer
answer = MsgBox("Er du sikker på at du vil slette opgave fra " + Range("C" & aCellR).value + "?", vbYesNo + vbQuestion, "Slet opgave")

If answer = vbYes Then
    'Slet rækken
    tbl_1.ListRows(currentR).Delete
    ActiveSheet.Preview.Text = ""
Else
    'Do nothing
End If

End Sub


Sub Knap7_Klik()

Set App = CreateObject("Outlook.Application")
Set NS = App.GetNamespace("MAPI")
NS.Logon

Dim aCellR, aCellRegionR, currentR As Integer
Set tbl_1 = ActiveSheet.ListObjects("Tabel1")

aCellR = ActiveCell.Row
aCellRegionR = ActiveCell.CurrentRegion.Row
currentR = aCellR - aCellRegionR


Set Msg = NS.GetItemFromID(Range("B" & aCellR).value)
Set repMsg = Msg.Reply

repMsg.Display


End Sub

Sub ViewEmail()

Set App = CreateObject("Outlook.Application")
Set NS = App.GetNamespace("MAPI")
NS.Logon

Dim aCellR, aCellRegionR, currentR As Integer
Set tbl_1 = ActiveSheet.ListObjects("Tabel1")

aCellR = ActiveCell.Row
aCellRegionR = ActiveCell.CurrentRegion.Row
currentR = aCellR - aCellRegionR


Set Msg = NS.GetItemFromID(Range("B" & aCellR).value)

Msg.Display


End Sub

