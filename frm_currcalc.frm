VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_currcalc 
   Caption         =   "Currency Converter"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11160
   OleObjectBlob   =   "frm_currcalc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_currcalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Populates both ComboBoxes from the hidden "tbl_currencylist" worksheet and table
Private Sub PopulateCurrencyDropdowns()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim desc As String
    
    cbo_currfrom.Clear
    cbo_currto.Clear
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("tbl_currencylist")
    If Not ws Is Nothing Then Set tbl = ws.ListObjects("tbl_currencylist")
    On Error GoTo 0
    
    If Not tbl Is Nothing Then
        If Not tbl.DataBodyRange Is Nothing Then
            For i = 1 To tbl.ListRows.Count
                desc = Trim(CStr(tbl.DataBodyRange(i, 1).Value))
                If desc <> "" Then
                    cbo_currfrom.AddItem desc
                    cbo_currto.AddItem desc
                End If
            Next i
        End If
        
        If cbo_currfrom.ListCount > 0 Then cbo_currfrom.ListIndex = 0
        If cbo_currto.ListCount > 0 Then cbo_currto.ListIndex = 0
    End If
End Sub

' Refreshes the external web query table inside "tbl_currencylist" and reloads dropdowns
Private Sub btn_refresh_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("tbl_currencylist")
    If Not ws Is Nothing Then Set tbl = ws.ListObjects("tbl_currencylist")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Or tbl Is Nothing Then
        MsgBox "The 'tbl_currencylist' worksheet or table could not be found.", vbCritical, "Refresh Error"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Refresh the query synchronously so code waits until download finishes
    tbl.QueryTable.Refresh BackgroundQuery:=False
    
    ' Standardize column headers and auto-fit column widths
    tbl.HeaderRowRange.Cells(1, 1).Value = "Currency Code with Country Name"
    tbl.HeaderRowRange.Cells(1, 2).Value = "Currency Code"
    ws.Columns("A:B").EntireColumn.AutoFit
    
    Call PopulateCurrencyDropdowns
    
    Application.ScreenUpdating = True
    MsgBox "Currency list successfully updated!", vbInformation, "Refresh Complete"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "An error occurred while refreshing: " & Err.Description, vbCritical, "Refresh Error"
End Sub

' Runs automatically when the UserForm opens: sets today's date and loads currencies
Private Sub UserForm_Initialize()
    txt_date.Text = Format(Date, "mm/dd/yyyy")
    Call PopulateCurrencyDropdowns
End Sub

' Closes the UserForm and clears it from memory
Private Sub btn_quit_Click()
    Unload Me
End Sub

' Validates user inputs, queries the API for the exchange rate, and displays the result
Private Sub btn_convert_Click()
    Dim amount As Double, convAmount As Double, rate As Double
    Dim fromCode As String, toCode As String
    Dim queryDate As Date
    
    If cbo_currfrom.ListIndex = -1 Or cbo_currto.ListIndex = -1 Then
        MsgBox "Please select both currencies.", vbCritical, "Input Error"
        Exit Sub
    End If
    
    If Trim(txt_amount.Text) = "" Or Not IsNumeric(txt_amount.Text) Then
        MsgBox "Amount must be a valid number.", vbCritical, "Input Error"
        Exit Sub
    End If
    
    amount = CDbl(txt_amount.Text)
    If amount <= 0 Then
        MsgBox "Amount must be greater than zero.", vbCritical, "Input Error"
        Exit Sub
    End If
    
    If Not TryParseDate(txt_date.Text, queryDate) Then
        MsgBox "Please enter a valid date in MM/DD/YYYY format.", vbCritical, "Input Error"
        Exit Sub
    End If
    
    If queryDate > Date Then
        MsgBox "Future dates are not available for exchange rates.", vbCritical, "Input Error"
        Exit Sub
    End If
    
    ' Extract the 3-letter ISO currency code (e.g., "EUR" from "EUR - Euro")
    fromCode = Trim(Left(cbo_currfrom.Text, 3))
    toCode = Trim(Left(cbo_currto.Text, 3))
    
    Application.ScreenUpdating = False
    rate = GetExchangeRate(fromCode, toCode, queryDate)
    Application.ScreenUpdating = True
    
    If rate > 0 Then
        convAmount = amount * rate
        txt_conv_amnt.Text = Format(convAmount, "0.000")
        MsgBox amount & " " & fromCode & " = " & Format(convAmount, "0.000") & " " & toCode & _
               vbCrLf & "Date: " & Format(queryDate, "mm/dd/yyyy"), vbInformation, "Conversion Successful"
    Else
        MsgBox "Could not retrieve exchange rate for the selected date/currencies.", vbCritical, "Query Failed"
    End If
End Sub

' Fetches 30 days of historical rates, builds a line chart, and activates the chart tab
Private Sub btn_plotlast30days_Click()
    Dim fromCode As String, toCode As String
    Dim endDate As Date, currDate As Date
    Dim wsPlot As Worksheet
    Dim chrtSheet As Chart
    Dim i As Integer
    Dim rate As Double
    
    If cbo_currfrom.ListIndex = -1 Or cbo_currto.ListIndex = -1 Then
        MsgBox "Please select both currencies.", vbCritical, "Input Error"
        Exit Sub
    End If

    If Not TryParseDate(txt_date.Text, endDate) Then
        MsgBox "Please enter a valid end date in MM/DD/YYYY format.", vbCritical, "Input Error"
        Exit Sub
    End If
    
    If endDate > Date Then
        MsgBox "Future dates are not available for exchange rates.", vbCritical, "Input Error"
        Exit Sub
    End If
    
    fromCode = Trim(Left(cbo_currfrom.Text, 3))
    toCode = Trim(Left(cbo_currto.Text, 3))
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo PlotErrorHandler
    
    ' Delete old chart if it exists and locate/create the data staging sheet
    On Error Resume Next
    ThisWorkbook.Charts("Plot last 30 days").Delete
    Set wsPlot = ThisWorkbook.Worksheets("30DayPlot")
    On Error GoTo PlotErrorHandler
    
    If wsPlot Is Nothing Then
        Set wsPlot = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsPlot.Name = "30DayPlot"
    End If
    
    ' Keep visible temporarily while writing data and binding the chart
    wsPlot.Visible = xlSheetVisible
    wsPlot.Cells.Clear
    wsPlot.Range("A1").Value = "Date"
    wsPlot.Range("B1").Value = "Exchange Rate (" & fromCode & " to " & toCode & ")"
    
    ' Populate rows 2 through 31 with 30 days of historical exchange rates
    For i = 29 To 0 Step -1
        currDate = DateAdd("d", -i, endDate)
        Application.StatusBar = "Querying historical data: " & Format(currDate, "mm/dd/yyyy") & " (" & (30 - i) & " of 30)..."
        
        rate = GetExchangeRate(fromCode, toCode, currDate)
        
        wsPlot.Range("A" & 31 - i).Value = currDate
        wsPlot.Range("B" & 31 - i).Value = rate
    Next i
    
    Application.StatusBar = False
    
    ' Create and format the dedicated chart sheet using the full 30-day range (A1:B31)
    Set chrtSheet = ThisWorkbook.Charts.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    With chrtSheet
        .Name = "Plot last 30 days"
        .ChartType = xlLine
        .PlotVisibleOnly = False ' Ensures chart still renders after wsPlot is hidden
        .SetSourceData Source:=wsPlot.Range("A1:B31")
        .ChartStyle = 235
        .HasTitle = True
        .ChartTitle.Text = "30 Day History: " & fromCode & " to " & toCode
        .SetElement (msoElementLegendBottom)
        .SetElement (msoElementPrimaryValueGridLinesNone)
        .Legend.Format.TextFrame2.TextRange.Font.Size = 10
        .Axes(xlCategory).TickLabels.NumberFormat = "mm/dd"
        .Tab.Color = 65535
        .Tab.TintAndShade = 0
    End With
    
    ' Hide staging data sheet and activate the new chart behind the modeless UserForm
    wsPlot.Visible = xlSheetVeryHidden
    chrtSheet.Activate
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "30-day plot generated successfully!" & vbCrLf & "You can now view and interact with the 'Plot last 30 days' tab behind this window.", vbInformation, "Plot Complete"
    Exit Sub

PlotErrorHandler:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "An error occurred while generating the plot: " & Err.Description, vbCritical, "Plot Error"
End Sub

