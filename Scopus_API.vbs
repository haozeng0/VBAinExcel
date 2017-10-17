Sub Scopus()
    Dim W As Worksheet: Set W = ActiveSheet
    Dim last As Integer: last = W.Range("A10000").End(xlUp).Row
    If last = 1 Then Exit Sub
    Dim ISSN As String
    Dim i As Integer

For i = 5 To last
        ISSN = W.Range("B" & i).Value
        Dim URL As String: URL = "https://api.elsevier.com/content/serial/title/issn/" + ISSN + "?apiKey=xxxxxxxxAPIxKEYxxxxxxxxxx&httpAccept=text%2Fxml"
        Dim Req As New XMLHTTP
        Req.Open "GET", URL, False
        Req.Send
         
        Dim Resp As New DOMDocument
        Resp.LoadXML Req.ResponseText
        
        Dim Title2 As IXMLDOMNode
        Dim Publisher2 As IXMLDOMNode
   '    Dim JournalURL As IXMLDOMNode
   '    Dim snipYear As IXMLDOMAttribute
        Dim snipScore As IXMLDOMNode
   '    Dim sjrYear As IXMLDOMAttribute
        Dim sjrScore As IXMLDOMNode
        Dim citeScoreYear As IXMLDOMNode
        Dim Tracker As IXMLDOMNode
        Dim TrackerYear As IXMLDOMNode
        
        
        For Each Title2 In Resp.SelectNodes("//dc:title")
            W.Range("J" & i).Value = Title2.Text
            On Error Resume Next
        Next Title2
        
        For Each Publisher2 In Resp.SelectNodes("//dc:publisher")
            W.Range("K" & i).Value = Publisher2.Text
            On Error Resume Next
        Next Publisher2
    
      '  For Each JournalURL In Resp.SelectNodes("//link ref='homepage' href=' '")
      '      W.Range("L" & i).Value = JournalURL.Text
      '      On Error Resume Next
      '  Next JournalURL
      
      ' For Each snipYear In Resp.Attributes("//SNIP/@year")
      '      W.Range("N" & i).Value = snipYear.Text
      '      On Error Resume Next
        ' Next snipYear

      For Each snipScore In Resp.SelectNodes("//SNIPList")
            W.Range("N" & i).Value = snipScore.Text
            On Error Resume Next
        Next snipScore
        
       ' For Each snipYear In Resp.Attributes("//SNIP/@year")
      '      W.Range("N" & i).Value = snipYear.Text
      '      On Error Resume Next
        ' Next snipYear
        
        For Each sjrScore In Resp.SelectNodes("//SJRList")
            W.Range("P" & i).Value = sjrScore.Text
            On Error Resume Next
        Next sjrScore
        
        For Each citeScoreYear In Resp.SelectNodes("//citeScoreCurrentMetricYear")
            W.Range("Q" & i).Value = citeScoreYear.Text
            On Error Resume Next
        Next citeScoreYear
        
        For Each citeScoreMetric In Resp.SelectNodes("//citeScoreCurrentMetric")
            W.Range("R" & i).Value = citeScoreMetric.Text
            On Error Resume Next
        Next citeScoreMetric
    
        For Each Tracker In Resp.SelectNodes("//citeScoreTracker")
            W.Range("S" & i).Value = Tracker.Text
            On Error Resume Next
        Next Tracker
        
        For Each TrackerYear In Resp.SelectNodes("//citeScoreTrackerYear")
            W.Range("T" & i).Value = TrackerYear.Text
            On Error Resume Next
        Next TrackerYear
        
    Application.Wait (Now + TimeValue("0:00:02"))
    
    Next i

End Sub
