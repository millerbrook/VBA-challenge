Sub Stock_Script():
      
    'Establish variables to count sheets and total number of stocks
    Dim Data_WS_Count As Integer
    Dim Sheet_Index As Integer
    Dim Total_Ticker_Row As Integer
    Total_Ticker_Row = 2 'Set to start at row 2)
   
    Data_WS_Count = ActiveWorkbook.Worksheets.Count 'Get count of number of sheets of data
    Sheets.Add(After:=Sheets(Data_WS_Count)).Name = "Totals_By_Year" 'Create a new Totals_By_Year sheet at end
       
    '***Looping through Sheets.
    For Sheet_Index = 1 To Data_WS_Count 'Data_WS_Count set to number of sheets w/ data (not incl Totals_By_Year)
       
       'Activate relevant sheet and save title in Sheet_Title variable (EVENTUALLY YEAR)
       Worksheets(Sheet_Index).Activate
       Dim Sheet_Title As String
       Sheet_Title = ActiveSheet.Name

        'PROBLEM: LOOPING THROUGH YEAR-BASED SHEETS AND CONCATENATING Totals_By_Year. Create a script that will loop through all the stocks for one year and output the following information.
        Dim Current_Row As Long 'Tracker and Loop variable for current row
        Current_Row = 2 'Initializes current row

        '***INDIVIDUAL STOCK TRACKING***

        Dim Current_Ticker As String
        Current_Ticker = Cells(2, 1).Value 'Initializes Ticker with first value in field
        
        Dim Ticker_Count As Integer 'Keeps track of the number of tickers
        Ticker_Count = 0 'Starts at 0, adds 1 at each break including final row [CHECK -- SHOULD THIS BE 0 to START B/C LAST ROW WILL STILL ADD 1?]

        Dim Initial_Open As Double 'Keeps track of Initial Opening Value for Current Stock
        Initial_Open = Cells(2, 3) 'Sets Open Value for First Stock

        Dim Daily_Change As Double 'Tracks daily change in a stock
        Daily_Change = 0

        Dim Sum_Daily_Change As Double 'Totals daily changes for a stock
        Sum_Daily_Change = 0 'Initializes at 0

        Dim Percent_Change As Double 'Calculates Percentage Changes for a stock
        Percent_Change = 0

        Dim Sum_Stock_Volume As LongLong 'Totals Stock Volume for a stock
        Sum_Stock_Volume = 0 'Initializes Total Stock Volume at 0

        '***SET UP FOR Totals_By_Year*** NOTE: Totals are calculated at breakpoints between tickers
        'Make Headers for Totals in Current Sheet
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Year"
        Range("K1").Value = "Total Change"
        Range("L1").Value = "Percent Change"
        Range("M1").Value = "Total Stock Volume"
        Range("N1").Value = "Initial Open"
        Range("N2").Value = Range("C2").Value
        Range("O1").Value = "Final Closing Value"
       'Format Columns
       Range("A1:O1").Font.Bold = True
       Range("A1:O1").Columns.AutoFit

        Dim Summary_Row As Integer 'Keeps track of which row to insert totals
        Summary_Row = 2

        Dim Total_Rows As Long 'Finds total rows on current sheet to set up loop
        Total_Rows = Cells(Rows.Count, 1).End(xlUp).Row + 1 'Adds 1 so loop considers first blank row
        
        '***SINGLE SHEET LOOP BEGINS***
        For Current_Row = 2 To Total_Rows 'uses Total_Rows to set correct loop length
            
            '***IF THIS ROW IS FOR A NEW STOCK, DEAL WITH PREVIOUS STOCK (TOTALS, ETC)
            If Cells(Current_Row, 1).Value <> Current_Ticker Then 'Check if this is a new ticker symbol

                    Ticker_Count = Ticker_Count + 1 'Note: this is global -- used in Totals_By_Year sheet, so doesn't reset

                    Range("I" & Summary_Row).Value = Current_Ticker 'Post Current Stock Ticker (for Ticker ending at previous row)
                
                    'Post Sheet Title (EVENTUALLY CHANGE TO YEAR) for record-keeping
                    Range("J" & Summary_Row).Value = Sheet_Title

                    'Range("K" & Summary_Row).Value = Sum_Daily_Change 'Posts Total Change (for Ticker ending at previous row)
                    
                    'Calculate and post percentage change
                    If Initial_Open = 0 Then
                        Percent_Change = 0 'Set percentage arbitrarily to 1000% if it isn't calculable
                    
                    Else
                        Percent_Change = (Cells(Current_Row - 1, 6).Value / Initial_Open) - 1
                        Range("L" & Summary_Row).Value = Percent_Change 'Post previous stock percent change
                        Range("L" & Summary_Row).NumberFormat = "0.00%" 'Set number format to percentage
                    
                    End If

                    Range("M" & Summary_Row).Value = Sum_Stock_Volume 'Posts Sum_Stock_Volume total for previous stock
                    
                    'Post Final Closing Value and Yearly Change for Previous stock
                    Range("O" & Summary_Row).Value = Cells(Current_Row - 1, 6).Value
                    Dim Stock_Yearly_Change As Double
                    Stock_Yearly_Change = Cells(Summary_Row, 15).Value - Cells(Summary_Row, 14).Value
                    Range("K" & Summary_Row).Value = Stock_Yearly_Change

            '***END DEALING WITH PREVIOUS STOCK. SEE IF THERE'S A NEW STOCK AND DEAL WITH IT***
                    'Check if loop has reached blank row. if so, exit for loop (b/c loop has completed page)
                    If Len(Cells(Current_Row, 1).Value) = 0 Then
                       Exit For
                    End If 'Exits loop on blank row

                    'Set Values for Next Stock (having checked that there IS a new stock)
                    Summary_Row = Summary_Row + 1 'Advance Summary_Row counter for next stock
                    Current_Ticker = Cells(Current_Row, 1).Value 'Selects next Stock Ticker
                    Initial_Open = Cells(Current_Row, 3).Value 'Sets Initial Open value for new stock
                    Range("N" & Summary_Row).Value = Initial_Open
                    Sum_Stock_Volume = 0 'Resets Sum of Stock Volume
                    Percent_Change = 0 'Resets Percent Change
                    Sum_Daily_Change = 0 'Resets Sum_Daily_Change
                          
            Else
                '***NO TICKER CHANGE DETECTED: FOR EACH ROW WHILE IN CURRENT TICKER***
                Sum_Stock_Volume = Sum_Stock_Volume + Cells(Current_Row, 7).Value 'Updates Sum of Stock Volume for Individual Stock
                
                Daily_Change = Cells(Current_Row, 6).Value - Cells(Current_Row, 3).Value 'Calculates Daily Change in Stock Value
                
                Sum_Daily_Change = Sum_Daily_Change + Daily_Change 'Updates Sum of Daily Change Value
            End If

        Next Current_Row
       
        'Select Range to put into Totals_By_Year Sheet
        Worksheets(Sheet_Index).Range("I2:O" & Ticker_Count + 1).Copy
        
        'Paste into Totals_By_Year Sheet
        ActiveSheet.Paste Destination:=Worksheets("Totals_By_Year").Range("A" & Total_Ticker_Row)

        'Increment Total_Ticker_Row for placement of next sheet's data in Totals_By_Year sheet
        Total_Ticker_Row = Total_Ticker_Row + Ticker_Count
    
    Next Sheet_Index

        'Format Totals_By_Year Sheet Headers and Percentages
        Sheets("Totals_By_Year").Activate
   
   '***Working with Totals_By_Year Sheet***
   '-------------------------------
    With ThisWorkbook.Sheets("Totals_By_Year")
            .Columns(4).Resize(.Rows.Count - 1, 1).Offset(1, 0).NumberFormat = "0.00%"
                
        Range("A1").Value = "Ticker"
        Range("B1").Value = "Year"
        Range("C1").Value = "Yearly Total Change"
        Range("D1").Value = "Yearly Percent Change"
        Range("E1").Value = "Yearly Total Stock Volume"
        Range("F1").Value = "Yearly Open Price"
        Range("G1").Value = "Yearly Closing Price"
        Range("A1:G1").Font.Bold = True
        Range("A1:G1").Columns.AutoFit
       

        '***Loop through Totals_By_Year Sheet to add: ***********************************************************************
        'Calculations and conditional formatting
        'Conditional formatting that will highlight positive change in green and negative change in red
        'and to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"***
        '************************************************************************************************************
        
        'Create variables to find required greatest numbers
        Dim Greatest_Pct_Increase As Single
        Greatest_Pct_Increase = Range("D2").Value
        Dim Greatest_Pct_Decrease As Single
        Greatest_Pct_Decrease = Range("D2").Value
        Dim Greatest_Total_Volume As Single
        Greatest_Total_Volume = Range("E2").Value

        'Sort Totals_By_Year worksheet
        With ActiveSheet.Sort
            .SortFields.Add Key:=Range("A1"), Order:=xlAscending
            .SortFields.Add Key:=Range("B1"), Order:=xlAscending
            .SetRange Range("A1:G" & Total_Ticker_Row)
            .Header = xlYes
            .Apply
       End With

       'Make By Year Calculations and Formatting
            For Summary_Row = 2 To Total_Ticker_Row
            
                'Check whether current Percent is biggest so far
                If Range("D" & Summary_Row).Value > Greatest_Pct_Increase Then
                    Greatest_Pct_Increase = Range("D" & Summary_Row).Value
                End If
                
                'Check whether current Percent is smallest so far
                If Range("D" & Summary_Row).Value < Greatest_Pct_Decrease Then
                    Greatest_Pct_Decrease = Range("D" & Summary_Row).Value
                End If

                'Check whether current volume is biggest so far
                If Range("E" & Summary_Row).Value > Greatest_Total_Volume Then
                    Greatest_Total_Volume = Range("E" & Summary_Row).Value
                End If

                'Apply Conditional Formatting
                If Range("C" & Summary_Row).Value > 0 Then
                    Range("C" & Summary_Row).Interior.ColorIndex = 4

                ElseIf Range("C" & Summary_Row).Value = 0 Then
                    Range("C" & Summary_Row).Interior.ColorIndex = 2
                
                Else
                    Range("C" & Summary_Row).Interior.ColorIndex = 3
                End If

        Next Summary_Row

        Range("H1").Value = "Greatest Percentage Increase"
        Range("I1").Value = Greatest_Pct_Increase
        Range("I1").NumberFormat = "0.00%"
        Range("H2").Value = "Greatest Percentage Decrease"
        Range("I2").Value = Greatest_Pct_Decrease
        Range("I2").NumberFormat = "0.00%"
        Range("H3").Value = "Greatest Total Volume"
        Range("I3").Value = Greatest_Total_Volume
        Range("H1:I3").Font.Bold = True
        Range("H1:I1").Columns.AutoFit
    'End Make By Year Calculations and Formatting
    End With

  'Create New Three_Year_Totals Sheet
    Sheets.Add(After:=Sheets("Totals_By_Year")).Name = "Three_Year_Totals" 'Create a new Three_Year_Totals sheet at end
  'Reset Variables and Repeat Earlier Parsing Individual Stocks Loop, But only for Totals_By_Year Sheet (to populate Three_Year_Totals sheet)
     'Activate relevant sheet and save title in Sheet_Title variable (EVENTUALLY YEAR)
     Worksheets("Totals_By_Year").Activate
  
  'Repeat Loop Through Totals_By_Year, but for Three_Year_Totals sheet
     'PROBLEM: LOOPING THROUGH YEAR-BASED SHEETS AND CONCATENATING Totals_By_Year. Create a script that will loop through all the stocks for one year and output the following information.
    ' Dim Current_Row As Long 'Tracker and Loop variable for current row
     Current_Row = 2 'Initializes current row

     '***INDIVIDUAL STOCK TRACKING***

     'Dim Current_Ticker As String
     Current_Ticker = Cells(2, 1).Value 'Initializes Ticker with first value in field
     
     'Dim Ticker_Count As Integer 'Keeps track of the number of tickers
     Ticker_Count = 1 'Starts at 0, adds 1 at each break including final row [CHECK -- SHOULD THIS BE 0 to START B/C LAST ROW WILL STILL ADD 1?]

     'Dim Initial_Open As Double 'Keeps track of Initial Opening Value for Current Stock
     Initial_Open = Cells(2, 3) 'Sets Open Value for First Stock

     'Dim Daily_Change As Double 'Tracks daily change in a stock
     Daily_Change = 0

     'Dim Sum_Daily_Change As Double 'Totals daily changes for a stock
     Sum_Daily_Change = 0 'Initializes at 0

     'Dim Percent_Change As Double 'Calculates Percentage Changes for a stock
     Percent_Change = 0

     'Dim Sum_Stock_Volume As LongLong 'Totals Stock Volume for a stock
     Sum_Stock_Volume = 0 'Initializes Total Stock Volume at 0

     '***SET UP FOR Three_Year_Totals*** NOTE: Totals are calculated at breakpoints between tickers
     'Make Headers for Totals in Current Sheet
     Range("AA1").Value = "Ticker"
     'Range("J1").Value = "Year" DON'T NEED FOR NEW SHEET (ARTEFACT FROM PRIOR LOOP)
     Range("AB1").Value = "Total Change"
     Range("AC1").Value = "Percent Change"
     Range("AD1").Value = "Total Stock Volume"
     Range("AE1").Value = "Initial Open"
     Range("AE2").Value = Range("F2").Value 'Grabs initial open for first stock
     Range("AF1").Value = "Final Closing Value"
    'Format Columns
    Range("AA1:AF1").Font.Bold = True
    Range("AA1:AF1").Columns.AutoFit

    ' Sets variables prior to going through Totals_By_Year to Extract Summary Info
     Summary_Row = 2
     Sum_Stock_Volume = 0

     'Dim Total_Rows As Long 'Finds total rows on current sheet to set up loop
     Total_Rows = Cells(Rows.Count, 1).End(xlUp).Row + 1 'Adds 1 so loop considers first blank row

     '***SINGLE SHEET LOOP THRU Totals_By_Year BEGINS***
     For Current_Row = 2 To Total_Rows 'uses Total_Rows to set correct loop length
         
         '***IF THIS ROW IS FOR A NEW STOCK, DEAL WITH PREVIOUS STOCK (TOTALS, ETC)
         If (Cells(Current_Row, 1).Value <> Current_Ticker) And (Current_Row > 2) Then 'Check if this is a new ticker symbol

                 Ticker_Count = Ticker_Count + 1 'Note: this is global -- used in Totals_By_Year sheet, so doesn't reset

                 Range("AA" & Summary_Row).Value = Current_Ticker 'Post Current Stock Ticker (for Ticker ending at previous row)
                                    
                 Range("AD" & Summary_Row).Value = Sum_Stock_Volume 'Posts Sum_Stock_Volume total for previous stock
                 
                 'Post Final Closing Value and Yearly Change for Previous stock
                 Range("AF" & Summary_Row).Value = Cells(Current_Row - 1, 7).Value
                 'Dim Stock_Yearly_Change As Double ARTEFACT FROM PRIOR LOOP
                 Stock_Yearly_Change = Cells(Summary_Row, 32).Value - Cells(Summary_Row, 31).Value
                 Range("AB" & Summary_Row).Value = Stock_Yearly_Change

         '***END DEALING WITH PREVIOUS STOCK. SEE IF THERE'S A NEW STOCK AND DEAL WITH IT***
                 'Check if loop has reached blank row. if so, exit for loop (b/c loop has completed page)
                 If Len(Cells(Current_Row, 1).Value) = 0 Then
                    Exit For
                 End If 'Exits loop on blank row

                 'Set Values for Next Stock (having checked that there IS a new stock)
                 Summary_Row = Summary_Row + 1 'Advance Summary_Row counter for next stock
                 Current_Ticker = Cells(Current_Row, 1).Value 'Selects next Stock Ticker
                 Initial_Open = Cells(Current_Row, 6).Value 'Sets Initial Open value for new stock
                 Range("AE" & Summary_Row).Value = Initial_Open
                 Sum_Stock_Volume = Cells(Current_Row, 5).Value 'Sets Sum of Stock Volume for New Stock
                 Percent_Change = 0 'Resets Percent Change
                 Sum_Daily_Change = 0 'Resets Sum_Daily_Change
           Else
             '***NO TICKER CHANGE DETECTED: FOR EACH ROW WHILE IN CURRENT TICKER***
             Sum_Stock_Volume = Sum_Stock_Volume + Cells(Current_Row, 5).Value 'Updates Sum of Stock Volume for Individual Stock
             
         End If

     Next Current_Row
    
    Total_Rows = Cells(Rows.Count, 27).End(xlUp).Row
    
    For Current_Row = 2 To Total_Rows
         'Calculate and post percentage change
         If Range("AE" & Current_Row) = 0 Then
            Range("AC" & Current_Row) = 0 'Set percentage arbitrarily to 0% if it isn't calculable
        
        Else
            Percent_Change = Range("AF" & Current_Row).Value / Range("AE" & Current_Row) - 1
            Range("AC" & Current_Row).Value = Percent_Change 'Post stock percent change
            Range("AC" & Current_Row).NumberFormat = "0.00%" 'Set number format to percentage
        
        End If
    Next Current_Row
    
     'Select Range to put into Totals_By_Year Sheet
     Worksheets("Totals_By_Year").Range("AA1:AF" & Ticker_Count + 1).Copy
     
     'Paste into Totals_By_Year Sheet
     ActiveSheet.Paste Destination:=Worksheets("Three_Year_Totals").Range("A1")

     'Format Totals_By_Year Sheet Headers and Percentages
     Sheets("Three_Year_Totals").Activate

'***Working with Three_Year_Totals Sheet***
'-------------------------------
    With ThisWorkbook.Sheets("Three_Year_Totals")
            .Columns(3).Resize(.Rows.Count - 1, 1).Offset(1, 0).NumberFormat = "0.00%"
                    
        Range("A1").Value = "Ticker"
        Range("B1").Value = "3 Year Total Change"
        Range("C1").Value = "3 Year Percent Change"
        Range("D1").Value = "3 Year Total Stock Volume"
        Range("E1").Value = "3 Year Open Price"
        Range("F1").Value = "3 Year Closing Price"
        Range("A1:F1").Font.Bold = True
        Range("A1:F1").Columns.AutoFit
            
            
        'Reset Variables to repeat conditional formatting, but now on Three_Year_Totals sheet
        Summary_Row = 0
        Greatest_Pct_Decrease = 0
        Greatest_Pct_Increase = 0
        Greatest_Total_Volume = 0

        Total_Rows = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Go through rows on Three_Year_Totals sheet
        For Summary_Row = 2 To Total_Rows
                    
                'Check whether current Percent is biggest so far
                If Range("C" & Summary_Row).Value > Greatest_Pct_Increase Then
                    Greatest_Pct_Increase = Range("C" & Summary_Row).Value
                End If
                
                'Check whether current Percent is smallest so far
                If Range("C" & Summary_Row).Value < Greatest_Pct_Decrease Then
                    Greatest_Pct_Decrease = Range("C" & Summary_Row).Value
                End If

                'Check whether current volume is biggest so far
                If Range("D" & Summary_Row).Value > Greatest_Total_Volume Then
                    Greatest_Total_Volume = Range("D" & Summary_Row).Value
                End If

                'Apply Conditional Formatting
                If Range("B" & Summary_Row).Value > 0 Then
                    Range("B" & Summary_Row).Interior.ColorIndex = 4

                ElseIf Range("B" & Summary_Row).Value = 0 Then
                    Range("B" & Summary_Row).Interior.ColorIndex = 2
                
                Else
                    Range("B" & Summary_Row).Interior.ColorIndex = 3
                End If

        Next Summary_Row
        
        'Post required values on Three_Year_Totals Sheet
        Range("H1").Value = "Greatest Percentage Increase"
        Range("I1").Value = Greatest_Pct_Increase
        Range("I1").NumberFormat = "0.00%"
        Range("H2").Value = "Greatest Percentage Decrease"
        Range("I2").Value = Greatest_Pct_Decrease
        Range("I2").NumberFormat = "0.00%"
        Range("H3").Value = "Greatest Total Volume"
        Range("I3").Value = Greatest_Total_Volume
        Range("H1:I3").Font.Bold = True
        Range("H1:I1").Columns.AutoFit

    End With
End Sub

