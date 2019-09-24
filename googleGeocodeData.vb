Function GoogleGeocode(address As String) As String
  Dim strAddress As String
  Dim strQuery As String
  Dim strLatitude As String
  Dim strLongitude As String

  strAddress = URLEncode(address)

  'Assemble the query string
  strQuery = "http://maps.googleapis.com/maps/api/geocode/xml?"
  strQuery = strQuery & "address=" & strAddress
  strQuery = strQuery & "&sensor=false"

  'define XML and HTTP components
  Dim googleResult As New MSXML2.DOMDocument
  Dim googleService As New MSXML2.XMLHTTP
  Dim oNodes As MSXML2.IXMLDOMNodeList
  Dim oNode As MSXML2.IXMLDOMNode

  'create HTTP request to query URL - make sure to have
  'that last "False" there for synchronous operation

  googleService.Open "GET", strQuery, False
  googleService.send
  googleResult.LoadXML (googleService.responseText)

  Set oNodes = googleResult.getElementsByTagName("geometry")

  If oNodes.Length = 1 Then
    For Each oNode In oNodes
      strLatitude = oNode.ChildNodes(0).ChildNodes(0).Text
      strLongitude = oNode.ChildNodes(0).ChildNodes(1).Text
      GoogleGeocode = strLatitude & "," & strLongitude
    Next oNode
  Else
    GoogleGeocode = "Not Found (try again, you may have done too many too fast)"
  End If
End Function