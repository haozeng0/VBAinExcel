Private Sub btnload_Click()
    Dim W As Worksheet: Set W = ActiveSheet
    Dim last As Integer: last = W.Range("A10000").End(xlUp).Row
    
    If last = 1 Then Exit Sub
    
    Dim ISSN As String
    
    
    Dim i As Integer
    
    For i = 2 To last
        ISSN = W.Range("D" & i).Value
        Dim URL As String: URL = "http://ulrichsweb.serialssolutions.com/api/xml/xxxxAPIxKEYxxxx/search?query=issn:" & ISSN

        Dim Req As New XMLHTTP
        Req.Open "GET", URL, False
        Req.Send
         
        Dim Resp As New DOMDocument
        Resp.LoadXML Req.ResponseText
        
        Dim LC As IXMLDOMNode
        Dim Subject As IXMLDOMNode
        
        For Each LC In Resp.SelectNodes("//UlrichTitle/LCNumber")
            W.Range("K" & i).Value = LC.Text
            On Error Resume Next
        Next LC

        For Each Subject In Resp.SelectNodes("//UlrichTitle/subject")
            W.Range("L" & i).Value = Subject.Text
            On Error Resume Next
            Debug.Print "test"

        Next Subject
      
       
    Next i
    
End Sub
