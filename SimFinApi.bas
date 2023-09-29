Attribute VB_Name = "SimFinApi"

Option Explicit

' execShell() function courtesy of Robert Knight via StackOverflow
' http://stackoverflow.com/questions/6136798/vba-shell-function-in-office-2011-for-mac
Private Declare PtrSafe Function popen Lib "libc.dylib" (ByVal command As String, ByVal mode As String) As LongPtr
Private Declare PtrSafe Function pclose Lib "libc.dylib" (ByVal file As LongPtr) As Long
Private Declare PtrSafe Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As Long
Private Declare PtrSafe Function feof Lib "libc.dylib" (ByVal file As LongPtr) As LongPtr

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

Function GetOperatingSystem() As String
    Dim os As String
    os = Application.OperatingSystem
    
    If InStr(1, os, "Windows") > 0 Then
        GetOperatingSystem = "Windows"
    ElseIf InStr(1, os, "Macintosh") > 0 Then
        GetOperatingSystem = "Mac"
    Else
        GetOperatingSystem = "Unknown"
    End If
End Function



Function execShell(command As String, Optional ByRef exitCode As Long) As String
    Dim file As LongPtr
    file = popen(command, "r")

    If file = 0 Then
        Exit Function
    End If

    While feof(file) = 0
        Dim chunk As String
        Dim read As Long
        chunk = Space(50)
        read = fread(chunk, 1, Len(chunk) - 1, file)
        If read > 0 Then
            chunk = Left$(chunk, read)
            execShell = execShell & chunk
        End If
    Wend

    exitCode = pclose(file)
End Function


Function SimFin(Ticker As String, Year As String, Period As String, Columname As String, Token As String, Optional Ttm As String = "false", Optional AsReported As String = "false") As Variant

    Dim JsonObject As Object
    Dim objRequest As Object
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    Dim os As String
    Dim output As Variant
    
    
    If IsMissing(AsReported) Then
        AsReported = "false"
    End If
    If IsMissing(Ttm) Then
        Ttm = "false"
    End If
    os = GetOperatingSystem()
    strUrl = "https://backend.simfin.com/api/v3/excel-plugin/statements?ticker=" + URLEncode(Ticker) + "&period=" + Period + "&fyear=" + Year + "&columnName=" + URLEncode(Columname) + "&asreported=" + AsReported + "&ttm=" + Ttm
    If os = "Windows" Then
        Set objRequest = CreateObject("MSXML2.ServerXMLHTTP")
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
       
    ElseIf os = "Mac" Then
        
        Dim curlCommand As String
        curlCommand = "curl -s -H 'Content-Type: application/json' -H 'Authorization: api-key " & Token & "' -o - """ & strUrl & """"
        
        strResponse = execShell(curlCommand)
       
    End If
     If IsNumeric(Trim(strResponse)) Then
            output = Trim(strResponse) * 1
        Else
            output = strResponse
        End If
        SimFin = output
        
End Function


Function SimFinPrices(Ticker As String, DateString As String, Columname As String, Token As String, Optional AsReported As String) As Variant

    Dim JsonObject As Object
    Dim objRequest As Object
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    Dim output As Variant

    If IsMissing(AsReported) Then
        AsReported = "false"
    End If

    DateString = Format(CDate(DateString), "yyyy-mm-dd")
    os = GetOperatingSystem()
    strUrl = "https://backend.simfin.com/api/v3/excel-plugin/prices?ticker=" + URLEncode(Ticker) + "&start=" + DateString + "&columnName=" + URLEncode(Columname) + "&asreported=" + AsReported
    If os = "Windows" Then
        Set objRequest = CreateObject("MSXML2.XMLHTTP")
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
    ElseIf os = "Mac" Then
        Dim curlCommand As String
        curlCommand = "curl -s -H 'Content-Type: application/json' -H 'Authorization: api-key " & Token & "' -o - """ & strUrl & """"
        strResponse = execShell(curlCommand)
    End If
    If IsNumeric(Trim(strResponse)) Then
        var1 = CDbl(Trim(strResponse))
        var2 = CDbl(Trim(Replace(strResponse, ".", ",")))
        t1 = Abs(var1)
        t2 = Abs(var2)
        If t2 < t1 Then
            output = t2
        Else
            output = t1
        End If
    Else
        output = strResponse
    End If
    SimFinPrices = output
    
    
End Function
