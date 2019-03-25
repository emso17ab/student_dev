Sub ArchiveMessage()

Dim OutlookApp As Outlook.Application
Dim OutlookNamespace As Namespace
Dim Folder As MAPIFolder
Dim OutlookMail As Variant

Set OutlookApp = New Outlook.Application
Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
Set Folder = OutlookNamespace.Folders("Mæglerservice 1. linje").Folders("Arkiv")

Dim aCellR, aCellRegionR, currentR As Integer
Set tbl_1 = ActiveSheet.ListObjects("Tabel1")

aCellR = ActiveCell.Row
aCellRegionR = ActiveCell.CurrentRegion.Row
currentR = aCellR - aCellRegionR

Set Msg = OutlookNamespace.GetItemFromID(Range("B" & aCellR).value)

Dim answer As Integer
answer = MsgBox("Er du sikker på at du vil slette opgave fra " + Range("C" & aCellR).value + "?", vbYesNo + vbQuestion, "Slet opgave")

If answer = vbYes Then
    'Slet rækken
    tbl_1.ListRows(currentR).Delete
    Msg.Move Folder
    ActiveSheet.Preview.Text = ""
Else
    'Do nothing
End If

End Sub

