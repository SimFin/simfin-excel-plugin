Attribute VB_Name = "SimFinApi"

Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False _
) As String

  Dim StringLen As Long: StringLen = Len(StringVal)

  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Char
        Case 32
          result(i) = Space
        Case 0 To 15
          result(i) = "%0" & Hex(CharCode)
        Case Else
          result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
End Function


Function SimFin(Ticker As String, Year As String, Period As String, Columname As String, Token As String, Optional Ttm As String, Optional AsReported As String) As Variant
    
    Dim JsonObject As Object
    Dim objRequest As Object
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    Dim output As Variant
    
    If IsMissing(AsReported) Then
        AsReported = "false"
    End If
    If IsMissing(Ttm) Then
        Ttm = "false"
    End If
    
    
    Set objRequest = CreateObject("MSXML2.XMLHTTP")
    strUrl = "https://backend.simfin.com/api/v3/excel-plugin/statements?ticker=" + URLEncode(Ticker) + "&period=" + Period + "&fyear=" + Year + "&columnName=" + URLEncode(Columname) + "&asreported=" + AsReported + "&ttm=" + Ttm
    blnAsync = True

    With objRequest
        .Open "GET", strUrl, blnAsync
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "api-key " + Token
        .Send
        'spin wheels whilst waiting for response
        While objRequest.readyState <> 4
            DoEvents
        Wend
        strResponse = .responseText
    End With
    If IsNumeric(Trim(strResponse)) Then
        output = Trim(strResponse) * 1
    Else
        output = strResponse
    End If
    SimFin = output
End Function

Function SimFinPrices(Ticker As String, Start As String, Columname As String, Token As String, Optional AsReported As String) As Variant
    
    Dim JsonObject As Object
    Dim objRequest As Object
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    Dim output As Variant
    
    If IsMissing(AsReported) Then
        AsReported = "false"
    End If
    
    Set objRequest = CreateObject("MSXML2.XMLHTTP")
    strUrl = "https://backend.simfin.com/api/v3/excel-plugin/prices?ticker=" + URLEncode(Ticker) + "&start=" + Start + "&columnName=" + URLEncode(Columname) + "&asreported=" + AsReported
    blnAsync = True
    
    With objRequest
        .Open "GET", strUrl, blnAsync
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "api-key " + Token
        .Send
        'spin wheels whilst waiting for response
        While objRequest.readyState <> 4
            DoEvents
        Wend
        strResponse = .responseText
    End With
    If IsNumeric(Trim(strResponse)) Then
        output = Trim(strResponse) * 1
    Else
        output = strResponse
    End If
    SimFinPrices = output
    
    
End Function

