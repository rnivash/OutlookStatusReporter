Attribute VB_Name = "TraneStatatusReport"
 Dim oMReadEmailContent As MReadEmailContent
 
Sub SendStatusEmail()
    
    Dim olApp As Outlook.Application, olNs As Outlook.NameSpace
    Dim olMail As Outlook.MailItem
    Dim eFolder As Outlook.Folder
    Dim i As Long, CurrentDate As Date
    
    Set oMReadEmailContent = New MReadEmailContent
    Set olApp = New Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    
    CurrentDate = Date - 1
   
    Set eFolder = olNs.GetDefaultFolder(olFolderInbox)
    
    For i = eFolder.Items.Count To 1 Step -1
    
            If TypeOf eFolder.Items(i) Is MailItem Then
            
                Set olMail = eFolder.Items(i)
                
                If InStr(olMail.Subject, "Huddle Data") > 0 And InStr(olMail.ReceivedTime, CurrentDate) > 0 Then
                
                    oMReadEmailContent.ReadMe (olMail.HTMLBody)
                   
                End If
                
                Set olMail = Nothing
                
            End If
            
    Next i
          
    CreateNewMessage
    
    Set oMReadEmailContent = Nothing
    Set olApp = Nothing
    Set olNs = Nothing
    'CurrentDate = Nothing
    Set eFolder = Nothing
    
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



