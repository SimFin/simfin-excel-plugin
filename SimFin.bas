Attribute VB_Name = "SimFin"
Function SimFin(Year As String, Week As String, Token As String) As String
    Dim JsonObject As Object
    Dim objRequest As Object
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    
    Set objRequest = CreateObject("MSXML2.XMLHTTP")
    strUrl = "https://api.collegefootballdata.com/games/players?year=" + Year + "&week=" + Week + "&seasonType=regular"
    blnAsync = True

    With objRequest
        .Open "GET", strUrl, blnAsync
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " + Token
        .Send
        'spin wheels whilst waiting for response
        While objRequest.readyState <> 4
            DoEvents
        Wend
        strResponse = .responseText
    End With
        Set JsonObject = JsonConverter.ParseJson(strResponse)
        MsgBox (JsonObject(1)("id"))
    SimFin = JsonObject(1)("id")
End Function



