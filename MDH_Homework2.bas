Attribute VB_Name = "Intermediate"
Sub Intermediate()
'Overview:
'   Create a script that will loop through one year of stock data for each run and return the total volume each stock had over that year
'   The "ticker" column is a unique three letter code to identify a public company
'   The "date" column is an eight digit date identifier in YYYYMMDD format. Only weekdays of a given year have data
'   The "open" column is the opening value of the stock
'   The "close" column is closing value of the stock
'   The "low" column is the lowest value of the stock on that day
'   The "high" column is the highest value of the stock on that day
'   The "vol" column the total amount of shares of that stock that changed hands that day
'References:
'   Storing the active sheet in an object:          http://codevba.com/excel/set_worksheet.htm#.XUN04ehKiUk
'   Create an auto dismissing message to the user:  https://www.tek-tips.com/viewthread.cfm?qid=1714553
'   Number formatting with VBA:                     https://docs.microsoft.com/en-us/office/vba/api/excel.range.numberformat


Dim thisTicker As String
Dim nextTicker As String
Dim lRow As Long
Dim tickSum As Double
Dim tickVol As Long
Dim tickCount As Long
Dim pctdone As Integer
Dim ws As Worksheet
Dim tickOpening As Double
Dim tickClosing As Double
Dim thisOpening As Double
Dim thisClosing As Double
Dim yearChange As Double

'#------------------------------------------------------------------------------
'I.Initialize
'#------------------------------------------------------------------------------

'Set active ws as the worksheet to use
Set ws = Application.ActiveSheet

'Clear Contents from I through L columns
ws.Range("I:L").ClearContents

'Intialize tickSum
tickSum = 0

'Initialize tickCount
tickCount = 1

'Initialize tickOpening
tickOpening = ws.Range("C2").Value

'Insert Headers into I through L columns
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Find the last non-blank cell in column A(1) to get row count
lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

'Warn the user that the program will be running for a while
Dim objShell, intButton
Set objShell = CreateObject("Wscript.Shell")
intButton = objShell.Popup("Please wait while Stock Exchange Data is Analyzed for: " + ws.Name, 2)

'Set formatting in columns I - L
ws.Range("J:J").NumberFormat = "0.00"
ws.Range("K:K").NumberFormat = "0.0%"


'#------------------------------------------------------------------------------
'II. Loop through tickers to get total volume, year-over-year change, and the year-over-year change as a percent of the year opening
'#------------------------------------------------------------------------------

For i = 2 To lRow
    
    'Get the current ticker ID
    thisTicker = ws.Cells(i, 1).Text
    
    'Get the next ticker ID
    nextTicker = ws.Cells(i + 1, 1).Text
    
    'Get the volume of that ticker for the day from column G
    tickVol = ws.Cells(i, Range("G1").Column).Value
    
    'Sum the volumes for the ticker
    tickSum = tickSum + tickVol

    'Get the opening value
    thisOpening = ws.Cells(i, Range("C1").Column).Value
    
    'Get the closing value
    thisClosing = ws.Cells(i, Range("F1").Column).Value
    
    'Check if this is the last instance of the ticker
    If ((thisTicker <> nextTicker) Or (i = lRow)) Then
    
        'Copy the ticker to column I
        ws.Cells(tickCount + 1, Range("I1").Column).Value = thisTicker
    
        'Calculate the year-over-year change
        yearChange = (thisClosing - tickOpening)
    
        'Copy the year-over-year change to column J
        ws.Cells(tickCount + 1, Range("J1").Column).Value = yearChange
        
        'Copy the year-over-year change in percent to column K
        'Watch out for divide by zero
        If tickOpening = 0 Then
            
            ws.Cells(tickCount + 1, Range("K1").Column).Value = "#VALUE"
        'If not dividing by zero:
        Else
        
            ws.Cells(tickCount + 1, Range("K1").Column).Value = yearChange / tickOpening
            
            'Set color based on positive or negative percent changes
            If yearChange <= 0 Then
                
                'If negative set to background color to red
                ws.Cells(tickCount + 1, Range("K1").Column).Interior.ColorIndex = 3
                
            Else
            
                'If positive set to green
                ws.Cells(tickCount + 1, Range("K1").Column).Interior.ColorIndex = 4
            
            End If
                
        
        End If
        
        'Copy the tickSum to column L
        ws.Cells(tickCount + 1, Range("L1").Column).Value = tickSum
    
        'Reinitialize tickSum
        tickSum = tickVol
        
        'Reinitialize tickOpening
        tickOpening = ws.Cells(i + 1, Range("C1").Column).Value
        
        'If so increase the count of unique tickers by one
        tickCount = tickCount + 1

    End If
            
Next i


End Sub

Sub Hard()
'#----------------------------
'# III. Loop throough results to provide information on the greatest % increase, greatest % decrease, and greatest total volume.
'#----------------------------

'initialize variables for largest and smallest so far as the first in the list
Dim maxIncrease As Double
Dim maxDecrease As Double
Dim maxVolume As Double
Dim incTicker As String
Dim decTicker As String
Dim volTicker As String
Dim thisChange As Double
Dim thisVolume As Double
Dim index As Integer
Dim lRow As Long
Dim ws As Worksheet

'Set active ws as the worksheet to use
Set ws = Application.ActiveSheet

'Clear contents from column M:Q
ws.Range("M:Q").ClearContents

'get number of results by looking for the last entry in column J
lRow = ws.Cells(ws.Rows.Count, Range("I1").Column).End(xlUp).Row

'Put labels in O2:O4
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Put labels in P1:Q1
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Format Q2:Q4
ws.Range("Q2:Q3").NumberFormat = "0.0%"
ws.Range("Q4").NumberFormat = "General"

'Bold this part of the report
ws.Range("O1:Q4").Font.Bold = True

'#--------------
'# Greatest percent increase
'#--------------

'Use the Max Excel function to get the max
maxIncrease = Application.WorksheetFunction.Max(ws.Range("K2:K" & lRow))

'Get the index of the max of range K using the Match excel function
index = Application.WorksheetFunction.Match(maxIncrease, ws.Range("K2:K" & lRow), 0)

'Get the ticker value using the Index excel function
incTicker = Application.WorksheetFunction.index(ws.Range("I2:I" & lRow), index)

'Populate Greatest increase fields
Range("Q2").Value = maxIncrease
Range("P2").Value = incTicker


'#--------------
'# Greatest percent decrease
'#--------------

'Use the MIN Excel function to get the min
maxDecrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & lRow))

'Get the index of the max of range K using the Match excel function
index = Application.WorksheetFunction.Match(maxDecrease, ws.Range("K2:K" & lRow), 0)

'Get the ticker value using the Index excel function
decTicker = Application.WorksheetFunction.index(ws.Range("I2:I" & lRow), index)

'Populate Greatest increase fields
Range("Q3").Value = maxDecrease
Range("P3").Value = decTicker


'#--------------
'# Greatest total volume
'#--------------

'Use the Max Excel function to get the max
maxVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & lRow))

'Get the index of the max of range L using the Match excel function
index = Application.WorksheetFunction.Match(maxVolume, ws.Range("L2:L" & lRow), 0)

'Get the ticker value using the Index excel function
volTicker = Application.WorksheetFunction.index(ws.Range("I2:I" & lRow), index)

'Populate Greatest increase fields
Range("Q4").Value = maxVolume
Range("P4").Value = volTicker

End Sub


