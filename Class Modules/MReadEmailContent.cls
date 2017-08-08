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

    Dim hd As New HTMLDocument
    hd.body.innerHTML = body
    
    Dim trs1 As HTMLTableRow, trs2 As HTMLTableRow
    
    Set trs1 = hd.getElementsByTagName("tr").Item(0)
    Set trs2 = hd.getElementsByTagName("tr").Item(1)
    
    
    Dim oWs As New WorkStatus
    
    oWs.ResourceName = trs2.cells.Item(0).innerHTML
    oWs.Application = trs2.cells.Item(1).innerHTML
    oWs.TaskName = trs2.cells.Item(2).innerHTML
    oWs.ActivityDesc = trs2.cells.Item(3).innerHTML
    oWs.StartedOn = trs2.cells.Item(4).innerHTML
    oWs.ExpectedCompletionOn = trs2.cells.Item(5).innerHTML
    oWs.Status = trs2.cells.Item(6).innerHTML
    oWs.Completion = trs2.cells.Item(7).innerHTML
    oWs.Remarks = "Remarks"
    oWorkStatusList.Add oWs
    
    Set oWs = Nothing
    Set hd = Nothing
    Set trs1 = Nothing
    Set trs2 = Nothing
    
    

End Sub

Public Function getXPathElement(sXPath As String, objElement As Object) As HTMLBaseElement
    Dim sXPathArray() As String
     
    Dim sNodeName As String
    Dim sNodeNameIndex As String
    Dim sRestOfXPath As String
    Dim lNodeIndex As Long
    Dim lCount As Long
     
     ' Split the xpath statement
    sXPathArray = Split(sXPath, "/")
    sNodeNameIndex = sXPathArray(1)
    If Not InStr(sNodeNameIndex, "[") > 0 Then
        sNodeName = sNodeNameIndex
        lNodeIndex = 1
    Else
        sXPathArray = Split(sNodeNameIndex, "[")
        sNodeName = sXPathArray(0)
        lNodeIndex = CLng(Left(sXPathArray(1), Len(sXPathArray(1)) - 1))
    End If
    sRestOfXPath = Right(sXPath, Len(sXPath) - (Len(sNodeNameIndex) + 1))
     
    Set getXPathElement = Nothing
    For lCount = 0 To objElement.ChildNodes().Length - 1
        If UCase(objElement.ChildNodes().Item(lCount).nodeName) = UCase(sNodeName) Then
            If lNodeIndex = 1 Then
                If sRestOfXPath = "" Then
                    Set getXPathElement = objElement.ChildNodes().Item(lCount)
                Else
                    Set getXPathElement = getXPathElement(sRestOfXPath, objElement.ChildNodes().Item(lCount))
                End If
            End If
            lNodeIndex = lNodeIndex - 1
        End If
    Next lCount
End Function

Public Function GetMeMailContent()
    

    GetMeMailContent = "<html>" & HeaderWithStyle & "<body>"
    
    GetMeMailContent = GetMeMailContent & "<table class='trtbl' >"
    
    GetMeMailContent = GetMeMailContent & GetTableRow
   
    
    GetMeMailContent = GetMeMailContent & "</table>"
    
    GetMeMailContent = GetMeMailContent & "</body></html>"

End Function

Private Function GetTableRow()



    GetTableRow = GetTableRow & "<tr ><td class='head'>Resource</td><td class='head'>Application</td><td class='head'>Activity</td><td class='head'>Activity Description</td><td class='head'>Start Date</td><td class='head'>End Date</td><td class='head'>Status</td><td class='head'>Completion %</td><td class='head'>Remarks</td></tr>"

    For i = oWorkStatusList.Count To 1 Step -1
    
        GetTableRow = GetTableRow & "<tr >"
        Set Item = oWorkStatusList.Item(i)
        GetTableRow = GetTableRow & "<td>" & Item.ResourceName & "</td>"
        GetTableRow = GetTableRow & "<td>" & Item.Application & "</td>"
        GetTableRow = GetTableRow & "<td>" & Item.TaskName & "</td>"
        GetTableRow = GetTableRow & "<td>" & Item.ActivityDesc & "</td>"
        GetTableRow = GetTableRow & "<td>" & Item.StartedOn & "</td>"
        GetTableRow = GetTableRow & "<td>" & Item.ExpectedCompletionOn & "</td>"
        GetTableRow = GetTableRow & "<td>" & Item.Status & "</td>"
        GetTableRow = GetTableRow & "<td>" & Item.Completion & "</td>"
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
    "table.trtbl tr td{ border:1px solid silver; margin:0px;padding:8px 10px 5px 10px;}" & _
    "table.trtbl tr td.head{ border:1px solid silver; margin:0px;padding:8px 10px 5px 10px;background-color:#237fd1;font-weight:400;color:#fff;}" & _
    " </style>" & _
  " </head>"

End Function
