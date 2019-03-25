Sub Auto_Open()

MsgBox "Welcome to OUTLOOK INCEPTION"
Call Refresh_proc

End Sub

Sub delay_10()

    Application.OnTime Now + TimeValue("00:05:00"), "Refresh_proc"
    
End Sub

Sub Refresh_proc()

'INFO: Update started
With Sheets("updating...")
.Visible = True
.Activate
End With
Application.Wait (Now + TimeValue("0:00:01"))

'Updating all data connections
ActiveWorkbook.RefreshAll

'Clearing the table so we can rebuild further down in the code
With Sheets(1).ListObjects("Tabel1")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With

'Integrating a connection with Outlook
Dim OutlookApp As Outlook.Application
Dim OutlookNamespace As Namespace
Dim Folder As MAPIFolder
Dim OutlookMail As Variant
Dim i As Integer

Set OutlookApp = New Outlook.Application
Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
Set Folder = OutlookNamespace.Folders("Mæglerservice 1. linje").Folders("Indbakke")

i = 1
'Getting every mailItem from 1.linje inbox
For Each OutlookMail In Folder.Items
        Range("eMail_subject").Offset(i, 0).value = OutlookMail.Subject
        Range("eMail_date").Offset(i, 0).value = OutlookMail.ReceivedTime
        Range("eMail_sender").Offset(i, 0).value = OutlookMail.SenderName
        Range("eMail_id").Offset(i, 0).value = OutlookMail.EntryID
        Range("eMail_unread").Offset(i, 0).value = OutlookMail.UnRead
        Range("eMail_att").Offset(i, 0).value = OutlookMail.Importance
        
i = i + 1
        
Next OutlookMail

Set Folder = Nothing
Set OutlookNamespace = Nothing
Set OutlookApp = Nothing

'Call delay_10

'INFO: Update finished
Application.Wait (Now + TimeValue("0:00:01"))
Sheets("updating...").Visible = False
Sheets(1).Activate


End Sub

