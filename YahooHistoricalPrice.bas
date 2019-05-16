Attribute VB_Name = "YahooHistoricalPrice"
Option Explicit

Public blnAbortProgress                     As Boolean
Public strSecurityCode                      As String

Public wsCurrentWorksheet                   As Worksheet

Public Const intDateCol = 1
Public Const intOpenPriceCol = 2
Public Const intHighPriceCol = 3
Public Const intLowPriceCol = 4
Public Const intClosePriceCol = 5
Public Const intAdjustedClosePriceCol = 6
Public Const intVolumneCol = 7
Public Const strHistoricalSheetName = "GetHistoricalPriceYahoo"



Private Function CheckToolWorksheets() As Boolean
'-------------------------------------------------------------------------------------------
' Purpose: to check if the core four worksheets exist
' If any one of these worksheets is missing, the tool may not work correctly.
' Otherwise the user/developer must change the corresponding VBA code.
'-------------------------------------------------------------------------------------------

Dim wsWorksheet                             As Worksheet
Dim response                                As Variant
Dim blnControlSheetFound                    As Boolean
Dim blnPriceSheetFound                      As Boolean
Dim blnPriceAdjSheetFound                   As Boolean
Dim blnReturnSheetFound                     As Boolean
Dim strPrompt                               As String

    blnControlSheetFound = False
    blnPriceSheetFound = False
    blnPriceAdjSheetFound = False
    blnReturnSheetFound = False
    
    For Each wsWorksheet In ThisWorkbook.Worksheets
        If wsWorksheet.Name = strHistoricalSheetName Then blnControlSheetFound = True
        If wsWorksheet.Name = "YHHistoricalPrices" Then blnPriceSheetFound = True
        If wsWorksheet.Name = "YHHistoricalPricesAdjusted" Then blnPriceAdjSheetFound = True
        If wsWorksheet.Name = "YHCalculatedReturns" Then blnReturnSheetFound = True
    Next
    If Not (blnControlSheetFound And blnPriceSheetFound And blnPriceAdjSheetFound And blnReturnSheetFound) Then
        strPrompt = "The following core worksheets are missing:" & vbCrLf & vbCrLf
        If Not blnControlSheetFound Then strPrompt = strPrompt + strHistoricalSheetName + vbCrLf
        If Not blnPriceSheetFound Then strPrompt = strPrompt + "YHHistoricalPrices" + vbCrLf
        If Not blnPriceAdjSheetFound Then strPrompt = strPrompt + "YHHistoricalPricesAdjusted " + vbCrLf
        If Not blnReturnSheetFound Then strPrompt = strPrompt + "YHCalculatedReturns" + vbCrLf
        strPrompt = strPrompt + vbCrLf + "Please do not rename the worksheet of this tool."
        MsgBox strPrompt, vbCritical + vbOKOnly, "Error"
        CheckToolWorksheets = False
    Else
        CheckToolWorksheets = True
    End If

End Function
Private Sub GetHistoricalData()
' Main Sub

Dim myRange                                 As Range
Dim wsWorksheet                             As Worksheet
Dim i                                       As Long
Dim intWorksheetStartRowJump                As Integer
Dim intErrorCount                           As Integer
Dim lngTotalErrorCount                      As Long
Dim lngTotalWarningCount                    As Long
Dim response                                As Variant
Dim dblSummaryPrice                         As Double
Dim dteSummaryDate                          As Date
Dim blnExtractError                         As Boolean
Dim blnWarningFound                         As Boolean
Dim blnWorksheetExist                       As Boolean
Dim strURL                                  As String
Dim strResponse                             As String
Dim lngFreezeRow                            As Long
    
    'check if core worksheets exist
    blnWorksheetExist = CheckToolWorksheets
    If Not blnWorksheetExist Then GoTo ExitEarly
    
    'clear data first
    Call ClearData(True)
    
    'Unfreeze
    If ThisWorkbook.Worksheets(strHistoricalSheetName) Is ActiveSheet Then
        With ThisWorkbook.Worksheets(strHistoricalSheetName)
            .Activate
            ActiveWindow.FreezePanes = False
        End With
    End If
    
    'Initialize
    blnAbortProgress = False
    Set wsWorksheet = ThisWorkbook.Worksheets(strHistoricalSheetName)
    Set myRange = wsWorksheet.Range("YHTickerInputHeading")
    intWorksheetStartRowJump = myRange.Row - 1
    Set myRange = wsWorksheet.Range("YHErrorsMessageHeading")
    myRange.EntireColumn.AutoFit
    wsWorksheet.Range("YHStatus").Value2 = "Initializing..."

    intErrorCount = 0
    lngTotalErrorCount = 0
    lngTotalWarningCount = 0
    
    Set myRange = wsWorksheet.Range("YHTickerInputHeading")
    i = 1
    Do While myRange.Offset(i, 0).Value <> ""
        strSecurityCode = Trim(myRange.Offset(i, 0).Value)
        wsWorksheet.Range("YHCurrentTicker").Value = strSecurityCode
        wsWorksheet.Range("YHStatus").Value = "Processing... " & strSecurityCode
        wsWorksheet.Range("YHProcess").Value = Format(CStr(i / (myRange.End(xlDown).Row - intWorksheetStartRowJump)), "0.00%")
        
        
        'Abort check
        DoEvents
        If blnAbortProgress Then
            response = MsgBox("Abort current process?" & vbCrLf & vbCrLf & _
                                    "Select 'Yes' to abort" & vbCrLf & _
                                    "Select 'No' to continue processing", vbQuestion + vbYesNo, "Confirm Abort?")
            If response = vbYes Then
                On Error GoTo 0
                GoTo ExitEarly
            End If
            'If vbNo
            blnAbortProgress = False
        End If

        'Get the Historical Data
        Call GetSecurityHistoricalData(dblSummaryPrice, dteSummaryDate, blnExtractError, blnWarningFound, strURL, strResponse)

        If Not blnExtractError Then
            myRange.Offset(i, 1).Value = Format(CStr(dblSummaryPrice), "#,##0.00")
            myRange.Offset(i, 2).Value = dteSummaryDate
            intErrorCount = 0
        Else
            Call ExtractErrorDetails(strResponse)
            myRange.Offset(i, 3).Value = "Extract Error"
            intErrorCount = intErrorCount + 1
            lngTotalErrorCount = lngTotalErrorCount + 1
            
            If intErrorCount > 10 Then
                response = MsgBox("The attempt to retrieve the Historical Data has failed at least 10 times consecutively for different security codes." _
                        & vbCrLf & vbCrLf & _
                        "The security codes may be invalid." & vbCrLf & _
                        "Your connection to the internet may no longer be available." & vbCrLf & _
                        "Your request may have been rejected due to limits on the source website." _
                        & vbCrLf & vbCrLf & _
                        "Select 'Yes' to continue to the next security." & vbCrLf & _
                        "Select 'No' to abort.", vbCritical + vbYesNo + vbDefaultButton2, "Error - Confirm to continue processing?")
                If response = vbNo Then GoTo ExitEarly
            End If
        End If
        
        'Set the Warning count
        If blnWarningFound Then
            myRange.Offset(i, 3).Value = "Data Warning"
            lngTotalWarningCount = lngTotalWarningCount + 1
        End If

        wsWorksheet.Range("YHStatus").Value = "Process completed for " & strSecurityCode
        wsWorksheet.Range("YHErrorCount").Value = lngTotalErrorCount
        wsWorksheet.Range("YHWarningCount").Value = lngTotalWarningCount
        
        
        'Abort check
        DoEvents
        If blnAbortProgress Then
            response = MsgBox("Abort current process?" & vbCrLf & vbCrLf & _
                                    "Select 'Yes' to abort" & vbCrLf & _
                                    "Select 'No' to continue processing", vbQuestion + vbYesNo, "Confirm Abort?")
            If response = vbYes Then
                On Error GoTo 0
                GoTo ExitEarly
            End If
            'If vbNo
            blnAbortProgress = False
        End If

        i = i + 1
    Loop
    
    Call CheckInconsistPeriods
    
    wsWorksheet.Range("YHStatus").Value = "Update Historical Data Complete."
    wsWorksheet.Range("YHProcess").Value = "100.00%"
    MsgBox "Update Historical Data Complete." & vbCrLf & vbCrLf _
        & "Total Errors: " & lngTotalErrorCount & vbCrLf _
        & "Total Warnings: " & lngTotalWarningCount _
        , vbInformation + vbOKOnly, "Info"
    
    'Freeze
    Application.ScreenUpdating = False
    wsWorksheet.Activate
    lngFreezeRow = ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHTickerInputHeading").Row + 1
    Rows(lngFreezeRow).Select
    ActiveWindow.FreezePanes = True
    wsWorksheet.Range("A1").Select
    wsCurrentWorksheet.Activate
    Application.ScreenUpdating = True

    Exit Sub
ExitEarly:
    With ThisWorkbook.Worksheets(strHistoricalSheetName)
        .Range("YHStatus").Value = "Operation Aborted"
    'Freeze
        Application.ScreenUpdating = False
        .Activate
        lngFreezeRow = ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHTickerInputHeading").Row + 1
        Rows(lngFreezeRow).Select
        ActiveWindow.FreezePanes = True
        .Range("A1").Select
        wsCurrentWorksheet.Activate
        Application.ScreenUpdating = True
    End With

End Sub
Private Sub GetSecurityHistoricalData(dblSummaryPrice As Double, dteSummaryDate As Date, _
                              blnExtractError As Boolean, blnWarningFound As Boolean, _
                              strYahooDataFileURL As String, strResponse As String)

'----------------------------------------------------------------------------------------------------------
' Purpose: return the most recent historical price trading date before the End Date
'          return URL and response details
'
' If no price can be found then $0.00 will be returned
' Note: By default, Excel VBA passes by reference.
'----------------------------------------------------------------------------------------------------------

Dim myRange                                    As Range
Dim dteStartDate                               As Date
Dim dteEndDate                                 As Date
Dim strStartDateUnix                           As String
Dim strEndDateUnix                             As String
Dim blnValidResponse                           As Boolean
Dim arrRows()                                  As String
Dim arrRow()                                   As String
Dim arrColumns()                               As String
Dim arrRowsAndColumns()                        As String
Dim i                                          As Long
Dim j                                          As Long
Dim myCopyRange                                As Range
Dim wsWorksheet                                As Worksheet
Dim strSaveToDirectory                         As String
Dim lngLastRow                                 As Long

    'Initialize
    dblSummaryPrice = 0
    blnExtractError = False
    blnValidResponse = True
    blnWarningFound = False
    With ThisWorkbook.Worksheets(strHistoricalSheetName)
        Select Case .Range("YHPeriod").Value
            Case "Daily"
                dteStartDate = .Range("YHBegDateDaily").Value
                dteEndDate = .Range("YHEndDateDaily").Value
            Case "Monthly"
                dteStartDate = .Range("YHBegDateMonthly").Value
                dteEndDate = .Range("YHEndDateMonthly").Value
            Case Else
                dteStartDate = .Range("YHBegDateDaily").Value
                dteEndDate = .Range("YHEndDateDaily").Value
        End Select
    End With
    dteSummaryDate = dteEndDate

    'Build URL request
    strStartDateUnix = strGetUnixDate(dteStartDate)
    strEndDateUnix = strGetUnixDate(dteEndDate)

    strYahooDataFileURL = strSetFinanceHistoryUrl(strStartDateUnix, strEndDateUnix)
    strResponse = strGetYahooFinanceDataRetry(strYahooDataFileURL, blnValidResponse)
    
    'Parse into matrix
    arrRows = Split(strResponse, vbLf)
    arrRow = Split(arrRows(0), ",")
    
    'Check Response
    If (Not blnValidResponse) Or arrRow(0) <> "Date" Then
        blnExtractError = True
        ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHCookie").Value = ""
        ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHCrumb").Value = ""
        Exit Sub
    End If

    For i = 0 To UBound(arrRows) - 1
        arrColumns = Split(arrRows(i), ",")
        
        ReDim Preserve arrRowsAndColumns(UBound(arrRows), UBound(arrColumns) + 1)
        For j = 0 To UBound(arrColumns)
            If j = 0 Then
                If i = 0 Then arrRowsAndColumns(i, j) = "Ticker"
                If i > 0 Then arrRowsAndColumns(i, j) = strSecurityCode
            End If
            arrRowsAndColumns(i, j + 1) = arrColumns(j)
        Next
    Next
    
    'Check for Nulls and other invalid data
    Call CheckDataWarnings(arrRowsAndColumns, blnWarningFound)

    'clear data area on the main worksheet
    With ThisWorkbook.Worksheets(strHistoricalSheetName)
        Set myRange = .Range("YHDataAreaHeadingStart").Offset(1, 0)
        lngLastRow = .Cells.SpecialCells(xlCellTypeLastCell).Row
        Set myRange = Range(myRange, .Cells(lngLastRow, .Range("YHDataAreaHeadingEnd").Column))
        myRange.Clear
    End With
    
    With ThisWorkbook.Worksheets(strHistoricalSheetName)
        'UNFINISHED
        'Export data to the current worksheet
        .Range("YHCurrentTicker").Value = strSecurityCode
        .Range("YHDataAreaUpHeadingTicker").Value = strSecurityCode
        Set myRange = .Range("YHDataAreaHeadingStart").Offset(1, 0)
        myRange.Resize(UBound(arrRowsAndColumns, 1), UBound(arrRowsAndColumns, 2) + 1).Value = arrRowsAndColumns
    
        'Reset the formats for Data Area - First the Date
        Set myRange = .Range("YHDataAreaDate")
        Set myRange = Range(myRange, myRange.End(xlDown))
        '04/25 fix change date format by VBA rather than set them manually
        myRange.NumberFormat = "m/d/yyyy"
        myRange.Value = myRange.Value2
        
        'Reset the formats for Data Area - Now the Prices
        Set myRange = .Range("YHDataAreaNumberHeadings")
        Set myRange = Range(myRange, myRange.End(xlDown))
        myRange.NumberFormat = "#,##0.00"
        myRange.Value = myRange.Value2
        
        'Reset formats for Data Area - Finally the Trade Volume
        Set myRange = .Range("YHDataAreaVolume")
        Set myRange = Range(myRange, myRange.End(xlDown))
        myRange.NumberFormat = "#,##0"
        myRange.Value = myRange.Value2
    
        Set myRange = .Range("YHDataAreaTicker")
        If myRange.Offset(1, 0).Value <> "" Then
            Set myRange = myRange.End(xlDown)
            Select Case .Range("YHDataType").Value
                Case "Prices": If IsNumeric(myRange.Offset(1, intClosePriceCol).Value) Then dblSummaryPrice = myRange.Offset(0, intClosePriceCol).Value
                Case "Dividends": If IsNumeric(myRange.Offset(1, intOpenPriceCol).Value) Then dblSummaryPrice = myRange.Offset(0, intOpenPriceCol).Value
                Case Else: If IsNumeric(myRange.Offset(1, intClosePriceCol).Value) Then dblSummaryPrice = myRange.Offset(0, intClosePriceCol).Value
            End Select
        Else
            dblSummaryPrice = 0
        End If
        
        If IsDate(myRange.Offset(0, intDateCol).Value) Then dteSummaryDate = myRange.Offset(0, intDateCol).Value
    End With
    
    Application.ScreenUpdating = False
    
    With ThisWorkbook.Worksheets(strHistoricalSheetName)
        Set myRange = .Range("YHDataAreaTicker")
        'Only create a worksheet or CSV file if the switch option is not "Default"
        'Only create a worksheet or CSV file if data has been returned
        If myRange.Offset(1, 0).Value <> "" Then Set myRange = Range(myRange, myRange.End(xlDown))
        If myRange.Rows.Count > 1 And .Range("YHPriceInfoSwitch").Value <> "Default" Then
            
            'Copy the range of Market Data that has been formatted ready for copying to the new Range on the Worksheet so that we don't need to format it again
            Set myCopyRange = .Range("YHDataAreaTicker")
            Set myCopyRange = Range(myCopyRange, .Range("YHDataAreaEnd"))
            Set myCopyRange = Range(myCopyRange, myCopyRange.End(xlDown))
            myCopyRange.Copy
        
            'Now after putting this data to the Control worksheet let's either create a Worksheet for the data or a CSV file in the same directory
            'depending on the selected Output option
            Set wsWorksheet = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
            wsWorksheet.Name = strGetWorksheetNextName(strSecurityCode)
            Set myRange = wsWorksheet.Range("A1")
            myRange.PasteSpecial
            myRange.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            wsWorksheet.Activate
            myRange.Select
            
            If .Range("YHPriceInfoSwitch").Value = .Range("YHOptCSV").Value Then
                    strSaveToDirectory = ThisWorkbook.Path
                    ThisWorkbook.Sheets(wsWorksheet.Name).Copy
                
                    ActiveWorkbook.SaveAs Filename:=strSaveToDirectory & "\" & wsWorksheet.Name & "_" & Format(Now(), "yyyymmdd\_hhmm") & ".csv", FileFormat:=xlCSV
                    ActiveWorkbook.Close savechanges:=False
                    ThisWorkbook.Activate
                                
                    Application.DisplayAlerts = False
                    wsWorksheet.Delete
                    Application.DisplayAlerts = True
            End If
            
        'UNFINISHED
        End If
        
        wsCurrentWorksheet.Activate
        
        Application.ScreenUpdating = True
        'Price and Returns
        Set myRange = .Range("YHDataAreaTicker")
        If myRange.Offset(1, 0).Value <> "" Then Set myRange = Range(myRange, myRange.End(xlDown))
        If myRange.Rows.Count > 1 Then
            If .Range("YHDataType").Value = "Prices" Then
                Call CopyCat
                Call FixEOMonth
            End If
        End If
    End With
    


End Sub
Private Function strSetFinanceHistoryUrl(strStartDateUnix As String, strEndDateUnix As String) As String
' This function will setup the URL that is used to collect the Historical Price
' It is equivalent to click "Historical Data" - "Download Data" on the website
' Crumb (which is necessary) is not included here yet

Dim strInterval                     As String
Dim strDataType                     As String
        
    'Set interval (daily, weekly, monthly) for the Extract
    Select Case ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHPeriod").Value
        Case "Daily": strInterval = "1d"
        Case "Weekly": strInterval = "1wk"
        Case "Monthly": strInterval = "1mo"
        Case Else: strInterval = "1d"
    End Select
    
    'Set the data type for the extract (Prices or Dividends)
    Select Case ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHDataType").Value
        Case "Prices": strDataType = "history"
        Case "Dividends": strDataType = "div"
        Case Else: strDataType = "history"
    End Select
    
    strSetFinanceHistoryUrl = "https://query1.finance.yahoo.com/v7/finance/download/" & strSecurityCode & _
        "?period1=" & strStartDateUnix & _
        "&period2=" & strEndDateUnix & _
        "&interval=" & strInterval & "&events=" & strDataType

End Function
Private Function strGetYahooFinanceDataRetry(strURL As String, blnValidResponse As Boolean) As String

Dim myErrorRange                                As Range
Dim strResult                                   As String
Dim arrRows()                                   As String
Dim arrRow()                                    As String
Dim i                                           As Integer
Dim strStartingURL                              As String
Dim blnForceRefresh                             As Boolean: blnForceRefresh = False

    'Initialize
    blnValidResponse = False
    strStartingURL = strURL
        
    'Loop through 2 times if it fails. If it fails it will get a new cookie and crumb
    For i = 1 To 2
        strURL = strStartingURL
        
        Select Case ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHExtractMethod").Value
            Case "WinHTTP": strResult = strGetYahooFinanceData(strURL, blnForceRefresh)
            Case "ServerXMLHTTP": strResult = strGetYahooFinanceDataServerXMLHTTP(strURL, blnForceRefresh)
        End Select
        
        'Test if it worked
        arrRows = Split(strResult, vbLf)
        arrRow = Split(arrRows(0), ",")
        If arrRow(0) = "Date" Then
            'Debug.Print "Number of Retrys to get Finance Data - " & i
            blnForceRefresh = False
            Exit For
        Else
            'Reset the crumb and cookie as they don't seem to work, this will mean a new set will be created
            blnForceRefresh = True
        End If
    Next i
    If blnForceRefresh = True Then
        blnValidResponse = False
    Else
        blnValidResponse = True
    End If
    strGetYahooFinanceDataRetry = strResult
        
End Function
Private Function strGetYahooFinanceData(strURL As String, Optional blnForceRefresh As Boolean = False) As String
'Reference: https://xlautomation.com.au/free-spreadsheets/yahoo-historical-price-extract
'This function will return the Yahoo Finance Data that has been requested via the URL
'The previously set Crumb and Cookie values will be re-used or reset

Dim strResult                                   As String
Dim objRequest                                  As WinHttp.WinHttpRequest
Dim strCrumb                                    As String
Dim strCookie                                   As String
  
    strGetYahooFinanceData = ""
    Call GetCrumbCookie(strCrumb, strCookie, blnForceRefresh)
    
    strURL = strURL + "&crumb=" + strCrumb
    
    Set objRequest = New WinHttp.WinHttpRequest
    With objRequest
        .Open "GET", strURL, False
        .SetRequestHeader "Cookie", strCookie
        '.setRequestHeader "Cache-Control", "no-cache, no-store, must-revalidate"           'Tried and did not make any  difference
        .Send
        .WaitForResponse (10)
        
        Call WriteResponseDetails("WinHTTP", "History", strURL, .ResponseText)
        strResult = .ResponseText
    End With
    
    strGetYahooFinanceData = strResult
        
End Function

Private Function strGetYahooFinanceDataServerXMLHTTP(strURL As String, Optional blnForceRefresh As Boolean = False) As String
'Reference: https://xlautomation.com.au/free-spreadsheets/yahoo-historical-price-extract
'A new function that extracts the data using a different method Server ServerXMLHTTP
'
'This function will return the Yahoo Finance Data that has been requested via the URL
'The previously set Crumb and Cookie values will be re-used or reset

Dim strResult                                   As String
Dim strCrumb                                    As String
Dim strCookie                                   As String
Dim objServerXMLHTTP                            As New MSXML2.ServerXMLHTTP60
    
    strGetYahooFinanceDataServerXMLHTTP = ""
    Call GetCrumbCookie(strCrumb, strCookie, blnForceRefresh)
    
    strURL = strURL + "&crumb=" + strCrumb
    
    With objServerXMLHTTP
        .Open "GET", strURL, False
        .SetRequestHeader "Cookie", strCookie
        .Send
        .WaitForResponse (10)
        
        Call WriteResponseDetails("ServerXMLHTTP", "History", strURL, .ResponseText)
        strResult = .ResponseText
    End With
    
    Set objServerXMLHTTP = Nothing
    
    strGetYahooFinanceDataServerXMLHTTP = strResult
        
End Function
Private Sub WriteResponseDetails(strMethod As String, strType As String, strURL As String, strText As String)
'Reference: https://xlautomation.com.au/free-spreadsheets/yahoo-historical-price-extract
'This routine is going to write the raw body and text information into the worksheet for review when there are problems

Dim myRange                         As Range
Dim i                               As Long

    If ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHDebugFlag").Value = "No" Then Exit Sub
    
    i = 1
    With ThisWorkbook.Worksheets(strHistoricalSheetName)
        Set myRange = .Range("YHResponseHeadingStart")
        If myRange.Offset(i, 0).Value <> "" Then i = myRange.End(xlDown).Row - .Range("YHResponseHeadingStart").Row + 1
    End With
    myRange.Offset(i, 0).Value = strSecurityCode
    myRange.Offset(i, 1).Value = strMethod
    myRange.Offset(i, 2).Value = strType
    myRange.Offset(i, 3).Value = strURL
    myRange.Offset(i, 3).WrapText = False
    myRange.Offset(i, 4).Value = strText
    myRange.Offset(i, 4).WrapText = False
    
End Sub
Sub ExtractErrorDetails(strResponse As String)
'Reference: https://xlautomation.com.au/free-spreadsheets/yahoo-historical-price-extract
'This routine is going to extract the Error details found for a particular Security Code
'It will list these separately in the worksheet and provide a "test" URL for the user to use in confirming the error

Dim myRange                             As Range
Dim i                                   As Integer
Dim arrErrorString()                    As String
Dim strErrorMessage                     As String
Dim intWorksheetStartRowJump            As Integer
    
    i = 1
    Set myRange = ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHErrorsTickerHeading")
    intWorksheetStartRowJump = myRange.Row - 1
    If myRange.Offset(1, 0).Value <> "" Then i = myRange.End(xlDown).Row - intWorksheetStartRowJump
    
    'First we will remove any Line Feed characters to get the error text in a single line
    strResponse = Replace(strResponse, vbLf, "")
    
    'Now split the Error Message and set the last value to be the Message we will display
    'Finally remove any " or } symbols
    arrErrorString = Split(strResponse, ":")
    strErrorMessage = arrErrorString(UBound(arrErrorString))
    strErrorMessage = Replace(strErrorMessage, Chr(34), "")
    strErrorMessage = LTrim(Replace(strErrorMessage, "}", ""))
    
    myRange.Offset(i, 0).Value = strSecurityCode
    myRange.Offset(i, 1).Value = strErrorMessage
    myRange.Offset(i, 2).Hyperlinks.Add _
        Anchor:=myRange.Offset(i, 2), _
        Address:="https://finance.yahoo.com/quote/" & strSecurityCode & "/history?p=" & strSecurityCode, _
        TextToDisplay:="Click to check " & strSecurityCode
    
    myRange.Offset(i, 1).EntireColumn.AutoFit

End Sub
Sub CheckDataWarnings(ByRef arrDataExtract() As String, blnWarningFound As Boolean)
'Reference: https://xlautomation.com.au/free-spreadsheets/yahoo-historical-price-extract
'This routine is going to go through the data that has been extracted and determining if there is any missing, invalid data

Dim i                       As Long
Dim j                       As Long
Dim blnNullDataFound        As Boolean
Dim blnInvalidDate          As Boolean
Dim blnInvalidNumber        As Boolean
Dim myRange                 As Range
Dim strWarningMessage       As String
Dim intWorksheetStartRowJump As Integer

    blnNullDataFound = False
    blnInvalidDate = False
    blnInvalidNumber = False
    blnWarningFound = False
    For i = 1 To UBound(arrDataExtract, 1) - 1
    
        For j = 0 To UBound(arrDataExtract, 2)
            If UCase(arrDataExtract(i, j)) = "NULL" Then blnNullDataFound = True
            
            Select Case j
                Case intDateCol
                    If Not IsDate(arrDataExtract(i, j)) Then blnInvalidDate = True
                Case intOpenPriceCol, intHighPriceCol, intLowPriceCol, intClosePriceCol, intAdjustedClosePriceCol, intVolumneCol
                    If Not IsNumeric(arrDataExtract(i, j)) Then blnInvalidNumber = True
            End Select
        Next
        
        'If we come across any invalid data then stop looking through the rest
        If blnNullDataFound Or blnInvalidDate Or blnInvalidNumber Then Exit For
    Next

    If Not (blnNullDataFound Or blnInvalidDate Or blnInvalidNumber) Then Exit Sub
    
    strWarningMessage = ""
    If blnNullDataFound Then strWarningMessage = "Nulls"
    If blnInvalidDate Then
        If strWarningMessage <> "" Then strWarningMessage = strWarningMessage & ", "
        strWarningMessage = strWarningMessage & "Invalid Dates"
    End If
    If blnInvalidNumber Then
        If strWarningMessage <> "" Then strWarningMessage = strWarningMessage & ", "
        strWarningMessage = strWarningMessage & "Invalid Numbers"
    End If
    
    'If we have found data in error then we will write a record for Warning
    blnWarningFound = True
    Set myRange = ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHErrorsTickerHeading")
    intWorksheetStartRowJump = myRange.Row - 1
    i = 1
    Set myRange = ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHErrorsTickerHeading")
    If myRange.Offset(1, 0).Value <> "" Then i = myRange.End(xlDown).Row - intWorksheetStartRowJump

    myRange.Offset(i, 0).Value = strSecurityCode
    myRange.Offset(i, 1).Value = "Warning: Data may contain - " & strWarningMessage
    myRange.Offset(i, 2).Hyperlinks.Add _
        Anchor:=myRange.Offset(i, 2), _
        Address:="https://finance.yahoo.com/quote/" & strSecurityCode & "/history?p=" & strSecurityCode, _
        TextToDisplay:="Click to check " & strSecurityCode
    
    Set myRange = ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHErrorsMessageHeading")
    myRange.EntireColumn.AutoFit
    
End Sub
Function strGetWorksheetNextName(strWorksheetName As String) As String
'This function will return the next available worksheet name
'If a Name already exist then the next available worksheet name is set by incrementing a count

Dim wsWorksheet                     As Worksheet
Dim i                               As Long
Dim strCurrentWorksheetName         As String
Dim intStartPosition                As Long
Dim intStopPosition                 As Long
Dim intMidLength                    As Long
Dim intMaxNumber                    As Long
Dim bFlag                           As Boolean


    strGetWorksheetNextName = strWorksheetName
    
    intMaxNumber = 0
    bFlag = False
    For Each wsWorksheet In ThisWorkbook.Worksheets
    
        'Remove any brackets that exist in the current Worksheet name
        strCurrentWorksheetName = wsWorksheet.Name
        intStartPosition = VBA.InStr(wsWorksheet.Name, "(")
        If intStartPosition > 0 Then strCurrentWorksheetName = VBA.Left(wsWorksheet.Name, intStartPosition - 1)
        
        intStopPosition = VBA.InStr(wsWorksheet.Name, ")")
        intMidLength = intStopPosition - intStartPosition - 1
        'lowercase problem
        If VBA.LCase(strCurrentWorksheetName) = VBA.LCase(strGetWorksheetNextName) Then
            If intStartPosition > 0 Then intMaxNumber = VBA.Mid(wsWorksheet.Name, intStartPosition + 1, intMidLength)
            bFlag = True
        End If
        
        
    Next
    
    If bFlag And intMaxNumber = 0 Then strGetWorksheetNextName = strGetWorksheetNextName & "(1)"
    If intMaxNumber > 0 Then strGetWorksheetNextName = strGetWorksheetNextName & "(" & intMaxNumber + 1 & ")"
    
End Function
Private Sub ClearData(Optional blnSilentmode As Boolean = False)

Dim myRange                       As Range
Dim wsWorksheet                   As Worksheet
Dim lngLastRow                    As Long
Dim response                      As Variant

    On Error GoTo RunTimeError
    'v2.0 fix: use clear rather than clearcontents to reduce used cells, thus shrinking the file size
    'v2.0 fix: sometimes last cell is the header. Bug fixed.
    
    If Not blnSilentmode Then
        response = MsgBox("This will clear all data in this tool. Are you sure?", _
            vbExclamation + vbYesNoCancel + vbDefaultButton1, "WARNING")
        If response = vbNo Or response = vbCancel Then Exit Sub
    End If
    
    'clear summary area
    Set wsWorksheet = ThisWorkbook.Sheets(strHistoricalSheetName)
    Set myRange = wsWorksheet.Range("YHSummaryPriceHeading").Offset(1, 0)
    lngLastRow = wsWorksheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    Set myRange = Range(myRange, wsWorksheet.Cells(lngLastRow, wsWorksheet.Range("YHErrorsURLHeading").Column))
    myRange.ClearContents
    
    'clear data area on the main worksheet
    Set myRange = wsWorksheet.Range("YHDataAreaHeadingStart").Offset(1, 0)
    Set myRange = Range(myRange, wsWorksheet.Cells(lngLastRow, wsWorksheet.Range("YHDataAreaHeadingEnd").Column))
    myRange.Clear
    wsWorksheet.Range("YHDataAreaUpHeadingTicker").Clear
    
    'clear response area on the main worksheet
    Set myRange = wsWorksheet.Range("YHResponseHeadingStart").Offset(1, 0)
    Set myRange = Range(myRange, wsWorksheet.Cells(lngLastRow, wsWorksheet.Range("YHResponseHeadingEnd").Column))
    myRange.Clear
    
    'clear price data in YHHistoricalPrices worksheet
    Set wsWorksheet = ThisWorkbook.Sheets("YHHistoricalPrices")
    Set myRange = wsWorksheet.Range("ResPriceHeadingStart")
    lngLastRow = wsWorksheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    If myRange.Row <= lngLastRow Then
        Set myRange = Range(myRange, wsWorksheet.Cells.SpecialCells(xlCellTypeLastCell))
        myRange.Clear
    End If
    
    'clear price data in YHHistoricalPricesAdjusted worksheet
    Set wsWorksheet = ThisWorkbook.Sheets("YHHistoricalPricesAdjusted")
    Set myRange = wsWorksheet.Range("ResPriceAdjHeadingStart")
    lngLastRow = wsWorksheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    If myRange.Row <= lngLastRow Then
        Set myRange = Range(myRange, wsWorksheet.Cells.SpecialCells(xlCellTypeLastCell))
        myRange.Clear
    End If
    
    'clear price data in YHCalculatedReturns worksheet
    Set wsWorksheet = ThisWorkbook.Sheets("YHCalculatedReturns")
    Set myRange = wsWorksheet.Range("ResReturnHeadingStart")
    lngLastRow = wsWorksheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    If myRange.Row <= lngLastRow Then
        Set myRange = Range(myRange, wsWorksheet.Cells.SpecialCells(xlCellTypeLastCell))
        myRange.Clear
    End If
    
    Exit Sub
RunTimeError:
    If Err.Number = 1004 Then
        MsgBox "Runtime Error. Please restart Excel. This is a bug of Excel.", vbCritical + vbOKOnly, "Error"
    End If
    Debug.Print "ClearData Failed!"
    blnAbortProgress = True
    
End Sub
Private Function strGetUnixDate(dteSetDate As Date) As String

'----------------------------------------------------------------------------------------------------
' This function return the "period" parameter in the Yahoo Finance URL
' The "period" parameter is a Unix Timestamp
' Currently, there is a 14400 (4:00 AM UTC or 0:00 AM EST) offset for the Yahoo website.
' On the days of Daylight Saving time, it is 18000 (5:00 AM UTC or 0:00 AM EST).
' But I have tested that using the same 14400 makes no difference to the result.
' Yahoo Finance may change their implementation so this value may change in the future.
' You may test it by comparing the result of Excel formula =(DATE(2018,3,20)-DATE(1970,1,1))*86400
' and the corresponding period parameter in the URL.
'----------------------------------------------------------------------------------------------------
    
    strGetUnixDate = (dteSetDate - DateValue("01/01/1970")) * 86400 + 14400

End Function
Private Sub CopyCat()

'before calling this function, make sure that there is non-empty data
'need strSecurityCode
'strSecurityCode = "AAPL"

Dim myRange              As Range
Dim myDestRange          As Range
Dim i                    As Integer
Dim intStartJump         As Integer
Dim dteDate              As Date
Dim intOffset            As Integer

    Application.ScreenUpdating = False
    'Date Column
    Set myRange = ThisWorkbook.Sheets(strHistoricalSheetName).Range("YHDataAreaDate")
    Set myRange = Range(myRange, myRange.End(xlDown))
    myRange.Copy
    
    Set myDestRange = ThisWorkbook.Sheets("YHHistoricalPrices").Range("ResPriceHeadingStart")
    myDestRange.PasteSpecial
    
    'Ticker heading
    If myDestRange.Offset(0, 1).Value = "" Then
        Set myDestRange = myDestRange.Offset(0, 1)
    Else
        Set myDestRange = myDestRange.End(xlToRight).Offset(0, 1)
    End If
    myDestRange.Value = strSecurityCode
    
    'copy closing price
    Set myRange = ThisWorkbook.Sheets(strHistoricalSheetName).Range("YHDataAreaClose").Offset(1, 0)
    Set myRange = Range(myRange, myRange.End(xlDown))
    myRange.Copy
    
    'paste closing price
    'v1.4 fix check number of periods for prices.
    'For example: TWTR. Twitter IPO in Nov 2013. Number of Prices available are not the same as others.
    If ThisWorkbook.Sheets(strHistoricalSheetName).Range("YHPeriod").Value = "Monthly" And myRange.Rows.Count < ThisWorkbook.Sheets(strHistoricalSheetName).Range("YHNumPeriod").Value Then
        Call AppendWarningMessage("Warning: Data may contain insufficient number of periods.")
        intOffset = ThisWorkbook.Sheets(strHistoricalSheetName).Range("YHNumPeriod").Value - myRange.Rows.Count
        Set myDestRange = myDestRange.Offset(1 + intOffset, 0)
        myDestRange.PasteSpecial
    Else
        Set myDestRange = myDestRange.Offset(1, 0)
        myDestRange.PasteSpecial
    End If
    
    'Now it is time for Adjusted price
    'Date Column
    Set myRange = ThisWorkbook.Sheets(strHistoricalSheetName).Range("YHDataAreaDate")
    Set myRange = Range(myRange, myRange.End(xlDown))
    myRange.Copy
    
    Set myDestRange = ThisWorkbook.Sheets("YHHistoricalPricesAdjusted").Range("ResPriceAdjHeadingStart")
    myDestRange.PasteSpecial
    
    'Ticker heading
    If myDestRange.Offset(0, 1).Value = "" Then
        Set myDestRange = myDestRange.Offset(0, 1)
    Else
        Set myDestRange = myDestRange.End(xlToRight).Offset(0, 1)
    End If
    myDestRange.Value = strSecurityCode
    
    'copy adjusted closing price
    Set myRange = ThisWorkbook.Sheets(strHistoricalSheetName).Range("YHDataAreaAdjClose").Offset(1, 0)
    Set myRange = Range(myRange, myRange.End(xlDown))
    myRange.Copy
    
    'paste adjusted closing price
    'v1.4 fix check number of periods for prices.
    'For example: TWTR. Twitter IPO in Nov 2013. Number of Prices available are not the same as others.
    If myRange.Rows.Count < ThisWorkbook.Sheets(strHistoricalSheetName).Range("YHNumPeriod").Value Then
        intOffset = ThisWorkbook.Sheets(strHistoricalSheetName).Range("YHNumPeriod").Value - myRange.Rows.Count
        Set myDestRange = myDestRange.Offset(1 + intOffset, 0)
        myDestRange.PasteSpecial
    Else
        Set myDestRange = myDestRange.Offset(1, 0)
        myDestRange.PasteSpecial
    End If
    
    'At last, the returns!
    'Date Column
    Set myRange = ThisWorkbook.Sheets(strHistoricalSheetName).Range("YHDataAreaDate")
    Set myRange = Range(myRange, myRange.End(xlDown))
    myRange.Copy
    
    Set myDestRange = ThisWorkbook.Sheets("YHCalculatedReturns").Range("ResReturnHeadingStart")
    myDestRange.PasteSpecial
    
    'Ticker heading
    If myDestRange.Offset(0, 1).Value = "" Then
        Set myDestRange = myDestRange.Offset(0, 1)
    Else
        Set myDestRange = myDestRange.End(xlToRight).Offset(0, 1)
    End If
    myDestRange.Value = strSecurityCode
    
    'FormulaR1C1
    With ThisWorkbook.Sheets("YHCalculatedReturns")
        i = .Range("ResReturnHeadingStart").End(xlDown).Row
        intStartJump = .Range("ResReturnHeadingStart").Row
        myDestRange.Offset(2, 0).FormulaR1C1 = "=IFERROR(YHHistoricalPricesAdjusted!RC/YHHistoricalPricesAdjusted!R[-1]C-1,  " & Chr(34) & Chr(34) & ")"
        Set myDestRange = Range(myDestRange.Offset(2, 0), myDestRange.Offset(i - intStartJump, 0))
        myDestRange.FillDown
        myDestRange.NumberFormat = "0.00%"
    End With
    
    ThisWorkbook.Sheets("YHHistoricalPrices").Activate
    ThisWorkbook.Sheets("YHHistoricalPrices").Range("A1").Select
    ThisWorkbook.Sheets("YHHistoricalPricesAdjusted").Activate
    ThisWorkbook.Sheets("YHHistoricalPricesAdjusted").Range("A1").Select
    ThisWorkbook.Sheets("YHCalculatedReturns").Activate
    ThisWorkbook.Sheets("YHCalculatedReturns").Range("A1").Select
    ThisWorkbook.Sheets(strHistoricalSheetName).Activate
    ThisWorkbook.Sheets(strHistoricalSheetName).Range("A1").Select
    
    'unfinished
    wsCurrentWorksheet.Activate
    Application.ScreenUpdating = True
    
    End Sub
Private Sub FixEOMonth()
'-----------------------------------------------------------------------------------------------------
' Purpose :set Date column to the end of a month if "monthly"
' In Yahoo Finance, for example, for monthly data of March 2017, the date column is always 03/01/2017 (except the current month).
' Opening price of a month is the opening price of the first trading day.
' High and Low is the highest and lowest point in the month.
' Closing price is the closing price of the last trading day of the month.
' For simplicity's sake, the tool use the end day of the month to replace the date column.
'------------------------------------------------------------------------------------------------------
    
Dim myDestRange          As Range
Dim dteDate              As Date
Dim i                    As Integer
    
    If ThisWorkbook.Sheets(strHistoricalSheetName).Range("YHPeriod").Value = "Monthly" Then
        Set myDestRange = ThisWorkbook.Sheets("YHHistoricalPrices").Range("ResPriceHeadingStart")
        i = 1
        Do While myDestRange.Offset(i, 0) <> ""
            dteDate = myDestRange.Offset(i, 0).Value
            dteDate = CDate(WorksheetFunction.EoMonth(dteDate, 0))
            myDestRange.Offset(i, 0).Value = dteDate
            i = i + 1
        Loop
        
        Set myDestRange = ThisWorkbook.Sheets("YHHistoricalPricesAdjusted").Range("ResPriceAdjHeadingStart")
        i = 1
        Do While myDestRange.Offset(i, 0) <> ""
            dteDate = myDestRange.Offset(i, 0).Value
            dteDate = CDate(WorksheetFunction.EoMonth(dteDate, 0))
            myDestRange.Offset(i, 0).Value = dteDate
            i = i + 1
        Loop
        
        Set myDestRange = ThisWorkbook.Sheets("YHCalculatedReturns").Range("ResReturnHeadingStart")
        i = 1
        Do While myDestRange.Offset(i, 0) <> ""
            dteDate = myDestRange.Offset(i, 0).Value
            dteDate = CDate(WorksheetFunction.EoMonth(dteDate, 0))
            myDestRange.Offset(i, 0).Value = dteDate
            i = i + 1
        Loop
    End If

End Sub
Private Sub AppendWarningMessage(strWarningMessage As String, Optional strTicker As String = "")
'append warning message to the list
Dim myRange                                As Range
Dim intWorksheetStartRowJump               As Integer
Dim i                                      As Integer

        Set myRange = ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHErrorsTickerHeading")
        intWorksheetStartRowJump = myRange.Row - 1
        i = 1
        Set myRange = ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHErrorsTickerHeading")
        If myRange.Offset(1, 0).Value <> "" Then i = myRange.End(xlDown).Row - intWorksheetStartRowJump
    
        If strTicker <> "" Then
            myRange.Offset(i, 0).Value = strTicker
        Else
            myRange.Offset(i, 0).Value = strSecurityCode
        End If
        myRange.Offset(i, 1).Value = strWarningMessage
        myRange.Offset(i, 2).Hyperlinks.Add _
            Anchor:=myRange.Offset(i, 2), _
            Address:="https://finance.yahoo.com/quote/" & strSecurityCode & "/history?p=" & strSecurityCode, _
            TextToDisplay:="Click to check " & strSecurityCode
        
        Set myRange = ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHErrorsMessageHeading")
        myRange.EntireColumn.AutoFit

End Sub
Private Sub CheckInconsistPeriods()

Dim myRange                                As Range
Dim myRangetemp                            As Range
Dim myTickerRange                          As Range
Dim lngLastRow                             As Long
Dim lngEmptyCells                          As Long
Dim i                                      As Long
Dim j                                      As Long
Dim blnTickerExist                         As Boolean


    With ThisWorkbook.Worksheets("YHHistoricalPrices")
        i = 1
        Set myRange = .Range("ResPriceHeadingStart")
        lngLastRow = myRange.End(xlDown).Row
        Do While myRange.Offset(0, i) <> ""
            Debug.Print myRange.Offset(0, i).Value
            blnTickerExist = False
            j = 1
            Set myRangetemp = Range(myRange.Offset(0, i), .Cells(lngLastRow, myRange.Offset(0, i).Column))
            lngEmptyCells = WorksheetFunction.CountBlank(myRangetemp)
            
            If lngEmptyCells <> 0 Then
                Set myTickerRange = ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHErrorsTickerHeading")
                Do While myTickerRange.Offset(j, 0) <> ""
                    If myTickerRange.Offset(j, 0).Value = myRange.Offset(0, i).Value Then
                       blnTickerExist = True
                       Exit Do
                    End If
                    j = j + 1
                Loop
                
                'unfinished
                'add check URL
                If (Not blnTickerExist) Then
                    Call AppendWarningMessage("Warning: Number of Periods is less than other tickers. Please manually check the dates for this stock.", _
                        myRange.Offset(0, i).Value)
                End If
            End If
            i = i + 1
        Loop
    End With
End Sub





Sub GetCrumbCookie(strCrumb As String, strCookie As String, Optional blnForceRefresh As Boolean)
'Reference: https://xlautomation.com.au/free-spreadsheets/yahoo-historical-price-extract
'This routine will return the Crumb and Cookie to be used in the call, if those values don't exist it will set new ones

    'Need to store and retrieve the cookie and crumb to save additional calls
    strCookie = ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHCookie").Value
    strCrumb = ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHCrumb").Value
    
    'If we dont have the cookie and crumb stored then go and get one and stored those values
    If blnForceRefresh Or strCookie = "" Or strCrumb = "" Then
        Select Case ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHExtractMethod").Value
            Case "WinHTTP": Call GetYahooRequest(strCrumb, strCookie)
            Case "ServerXMLHTTP": Call GetYahooRequestServerXMLHTTP(strCrumb, strCookie)
        End Select
        ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHCookie").Value = strCookie
        ThisWorkbook.Worksheets(strHistoricalSheetName).Range("YHCrumb").Value = strCrumb
    End If
    
End Sub
Sub GetYahooRequest(strCrumb As String, strCookie As String)
'Reference: https://xlautomation.com.au/free-spreadsheets/yahoo-historical-price-extract
'I fixed their mistake because https://finance.yahoo.com/lookup?s=AAPL is no longer working as expected.

'This routine will use a sample request to Yahoo to obtain a valid Cookie and Crumb

Dim strURL                      As String: strURL = "https://finance.yahoo.com/quote/AAPL/history?p=AAPL"
Dim objRequest                  As WinHttp.WinHttpRequest
    
    Set objRequest = New WinHttp.WinHttpRequest
        
    With objRequest
        .Open "GET", strURL, False
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        '.setRequestHeader "Cache-Control", "no-cache, no-store, must-revalidate"                   'Included but did not resolve issues
        .Send
        .WaitForResponse (10)
        
        Call WriteResponseDetails("WinHTTP", "Crumb", strURL, .ResponseText)
        strCrumb = strExtractCrumb(.ResponseText)
        strCookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)
    End With
    
End Sub
Sub GetYahooRequestServerXMLHTTP(strCrumb As String, strCookie As String)
'Reference: https://xlautomation.com.au/free-spreadsheets/yahoo-historical-price-extract
'I fixed their mistake because https://finance.yahoo.com/lookup?s=AAPL is no longer working as expected.

'A new routine that extracts the data using a different method Server ServerXMLHTTP
'This routine will use a sample request to Yahoo to obtain a valid Cookie and Crumb

Dim strURL                              As String: strURL = "https://finance.yahoo.com/quote/AAPL/history?p=AAPL"
Dim objServerXMLHTTP                    As New MSXML2.ServerXMLHTTP60
        
    With objServerXMLHTTP
        .Open "GET", strURL, False
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .Send
        .WaitForResponse (10)
        
        Call WriteResponseDetails("ServerXMLHTTP", "Crumb", strURL, .ResponseText)
        strCrumb = strExtractCrumb(.ResponseText)
        strCookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)
    End With
    
    Set objServerXMLHTTP = Nothing
    
End Sub
Private Function strExtractCrumb(strResponse As String) As String
'Reference: https://xlautomation.com.au/free-spreadsheets/yahoo-historical-price-extract
'This function will extract the crumb string from another string set ready for use in the extract of data from Yahoo
'Starting string    - "CrumbStore":{"crumb":"IaHeg5nioGe"}
'Extract            - IaHeg5nioGe

Dim strCrumbStart               As String
Dim x                           As Long
Dim strField                    As String
Dim strChar                     As String
Dim intCrumbStartPosition       As Long

    strExtractCrumb = ""
    strCrumbStart = Chr(34) & "CrumbStore" & Chr(34) & ":{" & Chr(34) & "crumb" & Chr(34) & ":" & Chr(34)
    
    If InStr(strResponse, strCrumbStart) = 0 Then Exit Function
    
    intCrumbStartPosition = InStr(strResponse, strCrumbStart)        'Set the starting position
    intCrumbStartPosition = intCrumbStartPosition + Len(strCrumbStart)  'Then jump to the end of the Start string
    
    strField = ""
    For x = intCrumbStartPosition To Len(strResponse)
        strChar = Mid(strResponse, x, 1)
        If strChar = Chr(34) Then
            strExtractCrumb = strField
            Exit For
        End If
        strField = strField + strChar
    Next
    
    'I fixed their \u002F mistake.
    If InStr(strExtractCrumb, "\u002F") <> 0 Then
        strExtractCrumb = Replace(strExtractCrumb, "\u002F", "/")
    End If
        
End Function



Private Sub RestoreFactorySettings()
    Dim response         As Variant
    
    response = MsgBox("Restore Factory Settings?", vbQuestion + vbYesNo, "Prompt")
    If response = vbNo Then Exit Sub

    With ThisWorkbook.Sheets(strHistoricalSheetName)
        .Range("YHPeriod").Value2 = .Range("YHFactoryPeriod").Value2
        .Range("YHInputArea1").Value2 = .Range("YHFactoryArea1").Value2
        .Range("YHInputArea2").Value2 = .Range("YHFactoryArea2").Value2
    End With
End Sub
Sub btnClearData()
    ThisWorkbook.Activate
    Set wsCurrentWorksheet = ActiveWorkbook.ActiveSheet
    Call ClearData(False)
End Sub
Sub btnAbort()
    ThisWorkbook.Activate
    Set wsCurrentWorksheet = ActiveWorkbook.ActiveSheet
    blnAbortProgress = True
End Sub
Sub btnRestoreFactorySettings()
    ThisWorkbook.Activate
    Set wsCurrentWorksheet = ActiveWorkbook.ActiveSheet
    Call RestoreFactorySettings
End Sub
Sub btnGetHistoricalData()
    ThisWorkbook.Activate
    Set wsCurrentWorksheet = ActiveWorkbook.ActiveSheet
    Call GetHistoricalData
End Sub

