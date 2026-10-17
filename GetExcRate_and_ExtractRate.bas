Attribute VB_Name = "GetExcRate_and_ExtractRate"
Option Explicit

' Launches the UserForm in Modeless mode so users can click Excel tabs while the form is open
Sub RunForm()
    frm_currcalc.Show vbModeless
End Sub

' Queries the currency API with automatic CDN fallback and returns the cross-rate
Function GetExchangeRate(fromCurr As String, toCurr As String, targetDate As Date) As Double
    Dim http As Object
    Dim url As String, fallbackUrl As String
    Dim response As String
    Dim fromRate As Double, toRate As Double
    Dim ws As Worksheet
    
    fromCurr = LCase(Trim(fromCurr))
    toCurr = LCase(Trim(toCurr))
    
    ' Shortcut: converting a currency to itself always yields 1
    If fromCurr = toCurr Then
        GetExchangeRate = 1
        Exit Function
    End If
    
    ' Locate or create the hidden background audit sheet without changing the active tab
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("HiddenQuery")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "HiddenQuery"
    End If
    ws.Visible = xlSheetVeryHidden
    On Error GoTo 0
    
    ws.Cells.Clear
    
    ' Select latest endpoint for today/yesterday, or historical date endpoint for older dates
    If targetDate >= Date - 1 Then
        url = "https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@latest/v1/currencies/usd.json"
        fallbackUrl = "https://latest.currency-api.pages.dev/v1/currencies/usd.json"
    Else
        url = "https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@" & Format(targetDate, "yyyy-mm-dd") & "/v1/currencies/usd.json"
        fallbackUrl = "https://" & Format(targetDate, "yyyy-mm-dd") & ".currency-api.pages.dev/v1/currencies/usd.json"
    End If
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    ' Prevent Excel from freezing on slow networks (Resolve, Connect, Send, Receive in ms)
    http.setTimeouts 5000, 5000, 10000, 10000
    
    On Error Resume Next
    http.Open "GET", url, False
    http.send
    
    ' Fallback to Cloudflare Pages mirror if jsDelivr fails or is blocked
    If Err.Number <> 0 Or http.Status <> 200 Then
        Err.Clear
        http.Open "GET", fallbackUrl, False
        http.send
    End If
    
    If Err.Number <> 0 Or http.Status <> 200 Then
        GetExchangeRate = 0
        Exit Function
    End If
    On Error GoTo 0
    
    response = http.responseText
    
    ' Log raw JSON payload (capped at Excel's 32,767 character cell limit)
    ws.Range("A1").Value = "API Response Data"
    ws.Range("A2").Value = Left(response, 32767)
    
    ' Extract rates relative to USD and compute cross-rate
    If fromCurr = "usd" Then
        fromRate = 1
    Else
        fromRate = ExtractRate(response, fromCurr)
    End If
    
    If toCurr = "usd" Then
        toRate = 1
    Else
        toRate = ExtractRate(response, toCurr)
    End If
    
    If fromRate > 0 And toRate > 0 Then
        GetExchangeRate = toRate / fromRate
    Else
        GetExchangeRate = 0
    End If
End Function

' Parses the nested "usd": { ... } JSON object to extract a specific currency's numeric rate
Function ExtractRate(json As String, currCode As String) As Double
    Dim usdObjPos As Long, startPos As Long, endPos As Long, bracePos As Long
    Dim searchStr As String, rateStr As String
    
    ' Step 1: Jump inside the "usd": { ... } object so top-level keys are ignored
    usdObjPos = InStr(1, json, """usd"":", vbTextCompare)
    If usdObjPos = 0 Then usdObjPos = 1
    
    ' Step 2: Locate the target currency key inside the nested object
    searchStr = """" & LCase(currCode) & """:"
    startPos = InStr(usdObjPos + 6, json, searchStr, vbTextCompare)
    
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        endPos = InStr(startPos, json, ",")
        bracePos = InStr(startPos, json, "}")
        
        ' Handle the final item in the JSON object (which ends with "}" instead of ",")
        If endPos = 0 Or (bracePos > 0 And bracePos < endPos) Then
            endPos = bracePos
        End If
        
        If endPos > startPos Then
            rateStr = Trim(Mid(json, startPos, endPos - startPos))
            ExtractRate = Val(rateStr)
        End If
    End If
End Function

' Safely parses MM/DD/YYYY or MM-DD-YYYY dates regardless of Windows regional settings
Function TryParseDate(ByVal dateStr As String, ByRef outDate As Date) As Boolean
    Dim parts() As String
    Dim cleanStr As String
    Dim m As Integer, d As Integer, y As Integer
    
    On Error GoTo ParseError
    cleanStr = Replace(Trim(dateStr), "-", "/")
    parts = Split(cleanStr, "/")
    
    If UBound(parts) = 2 Then
        m = CInt(parts(0))
        d = CInt(parts(1))
        y = CInt(parts(2))
        If m >= 1 And m <= 12 And d >= 1 And d <= 31 And y >= 1990 And y <= 2100 Then
            outDate = DateSerial(y, m, d)
            TryParseDate = True
            Exit Function
        End If
    End If
    
    If IsDate(dateStr) Then
        outDate = CDate(dateStr)
        TryParseDate = True
        Exit Function
    End If

ParseError:
    TryParseDate = False
End Function
