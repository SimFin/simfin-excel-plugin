Attribute VB_Name = "SimFinApi"
Function SimFin(Ticker As String, Year As String, Period As String, Columname As String, Token As String) As String
    
    Dim JsonObject As Object
    Dim objRequest As Object
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    
    
    Set objRequest = CreateObject("MSXML2.XMLHTTP")
    strUrl = "http://192.168.2.203:8081/api/v3/companies/statements/plugin?ticker=" + Ticker + "&period=" + Period + "&fyear=" + Year + "&end=2023-01-13&columnName=" + Columname + ""
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
        Set JsonObject = JsonConverter.ParseJson(strResponse)
    SimFin2 = JsonObject("value")
End Function
