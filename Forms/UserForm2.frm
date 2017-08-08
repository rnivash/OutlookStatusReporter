VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   7950
   ClientLeft      =   96
   ClientTop       =   432
   ClientWidth     =   10974
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    CreateNewMessage
End Sub

Private Sub UserForm_Click()

End Sub


Public Sub CreateNewMessage()
Dim objMsg As MailItem

Me.Hide


Set objMsg = Application.CreateItem(olMailItem)

 With objMsg
 
  .Subject = "Status Report"
  
  .BodyFormat = olFormatHTML
  .Importance = olImportanceNormal
  .Sensitivity = olNormal
  

  .Display
End With

Set objMsg = Nothing
End Sub
