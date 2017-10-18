
Sub Scopus()

    Dim W As Worksheet: Set W = ActiveSheet
    Dim last As Integer: last = W.Range("B1000").End(xlUp).Row
    Debug.Print last
    If last = 1 Then Exit Sub

    Dim ISSN As String
    Dim i As Integer

For i = 2 To last
        ISSN = W.Range("B" & i).Value
        Dim URL As String: URL = "https://api.elsevier.com/content/serial/title/issn/" + ISSN + "?apiKey=00cc42d32895faa073f62fee64b25acc&httpAccept=text%2Fxml"
        Dim Req As New XMLHTTP
        Req.Open "GET", URL, False
        Req.Send
         
        Dim Resp As New DOMDocument
        Resp.LoadXML Req.ResponseText
        
        Dim Title As IXMLDOMNode
        Dim EISSN As IXMLDOMNode
        Dim Publisher As IXMLDOMNode
        Dim AggregationType As IXMLDOMNode
        Dim OpenAccess As IXMLDOMNode
        Dim snipYear As IXMLDOMNode
        Dim snipScore As IXMLDOMNode
        Dim sjrYear As IXMLDOMNode
        Dim sjrScore As IXMLDOMNode
        Dim citeScoreYear As IXMLDOMNode
        Dim citeScoreMetric As IXMLDOMNode
        Dim TrackerYear As IXMLDOMNode
        Dim Tracker As IXMLDOMNode
    
        For Each Title In Resp.SelectNodes("//dc:title")
            W.Range("A" & i).Value = Title.Text
            On Error Resume Next
        Next Title
                        
        For Each EISSN In Resp.SelectNodes("//prism:eIssn")
            W.Range("C" & i).Value = EISSN.Text
            On Error Resume Next
        Next EISSN
        
        For Each Publisher In Resp.SelectNodes("//dc:publisher")
            W.Range("D" & i).Value = Publisher.Text
            On Error Resume Next
        Next Publisher
        
        For Each AggregationType In Resp.SelectNodes("//prism:aggregationType")
            W.Range("E" & i).Value = AggregationType.Text
            On Error Resume Next
        Next AggregationType
        
        For Each OpenAccess In Resp.SelectNodes("//openaccess")
            W.Range("F" & i).Value = OpenAccess.Text
            On Error Resume Next
        Next OpenAccess
    
        For Each snipYear In Resp.SelectNodes("//SNIPList/SNIP").NextNode.Attributes
            W.Range("G" & i).Value = snipYear.NodeValue
            On Error Resume Next
        Next snipYear
      
        For Each snipScore In Resp.SelectNodes("//SNIPList")
            W.Range("H" & i).Value = snipScore.Text
            On Error Resume Next
        Next snipScore
        
        For Each sjrYear In Resp.SelectNodes("//SJRList/SJR").NextNode.Attributes
            W.Range("I" & i).Value = sjrYear.Text
            On Error Resume Next
         Next sjrYear
        
        For Each sjrScore In Resp.SelectNodes("//SJRList")
            W.Range("J" & i).Value = sjrScore.Text
            On Error Resume Next
        Next sjrScore
        
        For Each citeScoreYear In Resp.SelectNodes("//citeScoreCurrentMetricYear")
            W.Range("K" & i).Value = citeScoreYear.Text
            On Error Resume Next
        Next citeScoreYear
        
        For Each citeScoreMetric In Resp.SelectNodes("//citeScoreCurrentMetric")
            W.Range("L" & i).Value = citeScoreMetric.Text
            On Error Resume Next
        Next citeScoreMetric
                
        For Each TrackerYear In Resp.SelectNodes("//citeScoreTrackerYear")
            W.Range("M" & i).Value = TrackerYear.Text
            On Error Resume Next
        Next TrackerYear
    
        For Each Tracker In Resp.SelectNodes("//citeScoreTracker")
            W.Range("N" & i).Value = Tracker.Text
            On Error Resume Next
        Next Tracker
        
    Application.Wait (Now + TimeValue("0:00:02"))
    
    Next i



End Sub

