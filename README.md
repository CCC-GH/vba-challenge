This Repository is for 02-VBA Homwwork
- Posted below (below "--"), is an Excel VBA script. Execute one time on a workbook and it will pass thru each worksheet-year generating a sumerized report
    note: Script will clear cells, format cells, and produce summerized report between columns H thru Q for earch year's worksheet
= Two files within repository: 
    1. "ccc_02-VBA_ScriptOnly_and_ScreenShot_Results" contains VBA pasted-script (1st tab), requirements, and screenshots (from #2 file) of 2016, 2015, 2016
    2. "ccc_02-VBA-Scripting_HW_Instructions_Resources_Multiple_year_stock_data" larger file, contains all data, stored script in Developer tab with executed results (screenshot #1 file).
------------------------------------------
'02-VBA-Homework by Carl Coffman
Sub stock_summary()
'Set pointer to first workbook, top-left, for content deletion warning
    Sheets(1).Select
    Application.Goto Range("A1"), 1
'Warn client columns H thru Q will be cleared of both formatting and content to make room for report
    If MsgBox("Warning: for " & Application.ThisWorkbook.Name & "'s Worksheets, columns H thru Q will be cleared - continue?", vbYesNo) = vbNo Then Exit Sub
    Dim Year_Tab As Worksheet
'Step through each worksheet within the workbook
    For Each Year_Tab In Worksheets
        Year_Tab.Select
        Application.Goto Range("A1"), 1
'Define variables
        Dim Ticker_Name As String
        Dim Year As Integer
        Dim First_Record As Boolean
        Dim Year_Open_Price As Double
        Dim Summary_Table_Row As Integer
        Dim Year_Stock_Volume As Double
        Dim Greatest_Stock_Value_Volume As Double
'Clear and set width and formating columns H thru Q for report
        Columns("I:Q").Clear
        Columns("H:Q").ColumnWidth = 12
        Columns("I:I").HorizontalAlignment = xlLeft
        Columns("I:I").ColumnWidth = 7
        Columns("J:L").HorizontalAlignment = xlRight
        Columns("K:L").ColumnWidth = 15
        Columns("J:J").NumberFormat = "#.##"
        Columns("N:P").HorizontalAlignment = xlLeft
        Columns("K:K").NumberFormat = "#.##%"
        Columns("I:P").Font.Bold = False
'Setup Column headers, title and format, for report
        Range("I" & 1).Value = "Ticker"
        Range("I" & 1).Font.Bold = True
        Range("I" & 1).Interior.ColorIndex = 15
        Range("J" & 1).Value = "Year Change"
        Range("J" & 1).Font.Bold = True
        Range("J" & 1).Interior.ColorIndex = 15
        Range("K" & 1).Value = "Year % Change"
        Range("K" & 1).Font.Bold = True
        Range("K" & 1).Interior.ColorIndex = 15
        Range("L" & 1).Value = "Year Volume"
        Range("L" & 1).Font.Bold = True
        Range("L" & 1).Interior.ColorIndex = 15
'Define counter and sum variables before looping throu records
        First_Record = True
        Greatest_Stock_Value_Perc_Yr_Inc = 0
        Greatest_Stock_Value_Perc_Yr_Dec = 0
        Greatest_Stock_Value_Year_Volume = 0
'Set summary table row to 2nd row (after title)
        Summary_Table_Row = 2
        Year_Open_Price = Cells(2, 3).Value
        Year_Stock_Volume = Cells(2, 7).Value
'Loop through all data rows
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
            If Year_Open_Price > 0 Then
'Check for last record within an individual ticker
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
'Since last record, start writing to summary table
                    Range("I" & Summary_Table_Row).Value = Cells(i, 1).Value
                    Range("J" & Summary_Table_Row).Value = Cells(i, 6).Value - Year_Open_Price
'If year's open price is lower(negative) than close price, set cell format to Red(3), otherwise (positive) Green(4)
                    If Cells(i, 6).Value - Year_Open_Price < 0 Then
                        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    Else
                        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    End If
'Calculate %price change between year's open and close price
                    Range("K" & Summary_Table_Row).Value = (Cells(i, 6).Value - Year_Open_Price) / Year_Open_Price
'Year's stock volume
                    Range("L" & Summary_Table_Row).Value = Year_Stock_Volume + Cells(i, 7).Value
'Check if greatest stock value increase for the year
                    If (Cells(i, 6).Value - Year_Open_Price) / Year_Open_Price > Greatest_Stock_Value_Perc_Yr_Inc Then
                        Greatest_Stock_Value_Perc_Yr_Inc = (Cells(i, 6).Value - Year_Open_Price) / Year_Open_Price
                        Greatest_Stock_Perc_Yr_Inc = Cells(i, 1).Value
                    End If
'Check if greatest stock value decrease for the year
                    If (Cells(i, 6).Value - Year_Open_Price) / Year_Open_Price < Greatest_Stock_Value_Perc_Yr_Dec Then
                        Greatest_Stock_Value_Perc_Yr_Dec = (Cells(i, 6).Value - Year_Open_Price) / Year_Open_Price
                        Greatest_Stock_Perc_Yr_Dec = Cells(i, 1).Value
                    End If
'Check if greatest stock volume for the year
                    If (Year_Stock_Volume + Cells(i, 7).Value) > Greatest_Stock_Value_Year_Volume Then
                        Greatest_Stock_Value_Year_Volume = Year_Stock_Volume + Cells(i, 7).Value
                        Greatest_Stock_Year_Volume = Cells(i, 1).Value
                    End If
'Move the Summary table row-pointer down 1 row
                    Summary_Table_Row = Summary_Table_Row + 1
                    First_Record = True
                Else
'If first record of an individual ticker - set year open price and reset stock-volume sum to first record
                    If First_Record = True Then
                        Year_Open_Price = Cells(i, 3).Value
                        First_Record = False
                        Year_Stock_Volume = Cells(i, 7).Value
                    Else
'Not first record of an individual ticker
                       Year_Stock_Volume = Year_Stock_Volume + Cells(i, 7).Value
'Reset year's stock volume
                    End If
                End If
            End If
        Next i
'Set formating for Summary
        Columns("N:N").ColumnWidth = 20
        Columns("O:O").ColumnWidth = 7
        Range("N" & 1).Value = Left(Cells(2, 2).Value, 4) + " Year"
        Range("N" & 1).Font.Bold = True
        Range("N" & 1).Interior.ColorIndex = 15
        Range("O" & 1).Value = "Ticker"
        Range("O" & 1).Font.Bold = True
        Range("O" & 1).Interior.ColorIndex = 15
        Range("P" & 1).Value = "Value"
        Range("P" & 1).Font.Bold = True
        Range("P" & 1).Interior.ColorIndex = 15
        Range("N" & 2).Value = "Greatest % Increase"
'Write out stored values
        Range("O" & 2).Value = Greatest_Stock_Perc_Yr_Inc
        Range("P" & 2).Value = Greatest_Stock_Value_Perc_Yr_Inc
        Range("P" & 2).NumberFormat = "#.##%"
        Range("N" & 3).Value = "Greatest % Decrease"
        Range("O" & 3).Value = Greatest_Stock_Perc_Yr_Dec
        Range("P" & 3).Value = Greatest_Stock_Value_Perc_Yr_Dec
        Range("P" & 3).NumberFormat = "#.##%"
        Range("N" & 4).Value = "Greatest Total Volume"
        Range("O" & 4).Value = Greatest_Stock_Year_Volume
        Range("P" & 4).Value = Greatest_Stock_Value_Year_Volume
    Next
End Sub
