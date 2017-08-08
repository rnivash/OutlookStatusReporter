VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MReadEmailContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private oWorkStatusList As New Collection


Public Sub ReadMe(body As String)

    Dim oWs As New WorkStatus
    
    oWs.TaskName = "Test name"
    oWs.Completion = "90%"
    oWs.StartedOn = Date
    oWs.ExpectedCompletionOn = Date
    oWs.Remarks = "Remarks"
    
    oWorkStatusList.Add oWs

End Sub

Public Function GetMeMailContent()
    

    GetMeMailContent = "<html>" & HeaderWithStyle & "<body>"
    
    GetMeMailContent = GetMeMailContent & "<table class='trtbl' >"
    
    GetMeMailContent = GetMeMailContent & GetTableRow
   
    
    GetMeMailContent = GetMeMailContent & "</table>"
    
    GetMeMailContent = GetMeMailContent & "</body></html>"

End Function

Private Function GetTableRow()

    

    For i = oWorkStatusList.Count To 1 Step -1
    
        GetTableRow = GetTableRow & "<tr >"
        
        Set Item = oWorkStatusList.Item(i)
    
        GetTableRow = GetTableRow & "<td>" & Item.TaskName & "</td>"
        GetTableRow = GetTableRow & "<td>" & Item.Completion & "</td>"
        GetTableRow = GetTableRow & "<td>" & Item.StartedOn & "</td>"
        GetTableRow = GetTableRow & "<td>" & Item.ExpectedCompletionOn & "</td>"
        GetTableRow = GetTableRow & "<td>" & Item.Remarks & "</td>"
        
        GetTableRow = GetTableRow & "</tr>"
    
    Next i
    
    

End Function

Private Function HeaderWithStyle()

    HeaderWithStyle = "<head>"
    
    HeaderWithStyle = HeaderWithStyle & "<meta name=""viewport"" content=""width=device-width"">" & _
    "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">" & _
    "<style type=""text/css"">" & _
    "table.trtbl { border-collapse: collapse;}" & _
    "table.trtbl tr td{ border:1px solid silver; margin:0px;padding:5px;}" & _
    " </style>" & _
  " </head>"

End Function
