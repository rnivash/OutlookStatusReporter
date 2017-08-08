Attribute VB_Name = "TraneStatatusReport"
 Dim oMReadEmailContent As MReadEmailContent
 
Sub SendStatusEmail()

    Set oMReadEmailContent = New MReadEmailContent
    
    Dim olApp As Outlook.Application, olNs As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder, olMail As Outlook.MailItem
    Dim eFolder As Outlook.Folder '~~> additional declaration
    Dim i As Long
    Dim x As Date
    Dim lrow As Long
    Dim tstmsg As String
    
    Set olApp = New Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    x = Date
    
    x = x - 1
   
    Set eFolder = olNs.GetDefaultFolder(olFolderInbox)
    
    For i = eFolder.Items.Count To 1 Step -1
            If TypeOf eFolder.Items(i) Is MailItem Then
                Set olMail = eFolder.Items(i)
                'And InStr(olMail.ReceivedTime, x) > 0
                If InStr(olMail.Subject, "wip") > 0 Then
                    oMReadEmailContent.ReadMe olMail.body
                   
                   
                End If
            End If
    Next i
  
    CreateNewMessage
    
End Sub

Public Sub CreateNewMessage()

    Dim objMsg As MailItem
    
    Set objMsg = Application.CreateItem(olMailItem)

    With objMsg
        .Subject = "Status Report"
        .BodyFormat = olFormatHTML
        .HTMLBody = oMReadEmailContent.GetMeMailContent
        .Importance = olImportanceNormal
        .Sensitivity = olNormal
        .Display
    End With

    Set objMsg = Nothing
    
End Sub



