Sub replyEmail()

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

