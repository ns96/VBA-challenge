Option Explicit

'Keep track of max row for the sheets
Dim finalRow As Long

'Declare some constants for colors
Const myRed As Integer = 3
Const myGreen As Integer = 4

Sub ProcessSheets()
    Dim wsCount As Integer
    wsCount = ActiveWorkbook.Worksheets.Count
    
    'Loop through sheet and return the name
    Dim wsName As String
    Dim currentWS As Worksheet
    
    Dim i As Integer
    For i = 1 To wsCount
        'Get the worksheet object
        Set currentWS = ActiveWorkbook.Worksheets(i)
        
        'Get the name and the year
        wsName = currentWS.Name
        
        'Grab the location of final row of sheet
        finalRow = currentWS.Cells(Rows.Count, 1).End(xlUp).Row
        
        Debug.Print ("Processing: " & wsName & " Rows: " & finalRow)
        
        'Process stocks on this sheets
        ProcessStocks currentWS
        
        Debug.Print ("Finished Processing: " & wsName)
    Next i
End Sub

Sub ProcessStocks(currentWS As Worksheet)
    Dim i, j As Long
    Dim volume, volumeTotal, openPrice, closePrice As Double
    Dim yearOpenPrice, yearClosePrice, yearChange, percentChange As Double
    Dim currentTicker, nextTicker As String
    
    'Variables to keep track of the max yearly and percent change
    Dim maxIncreaseTicker, maxDecreaseTicker, maxVolumeTicker As String
    Dim maxIncrease, maxDecrease, maxVolume As Double
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0
    
    'Initialize some values
    yearOpenPrice = -1
    volumeTotal = 0
    
    'keep track of the row for a particular stock card
    j = 2
    
    'Add the summation table header start at "I"
    currentWS.Range("I1").Value = "Ticker"
    currentWS.Range("J1").Value = "Yearly Change"
    currentWS.Range("K1").Value = "Percentage Change"
    currentWS.Range("L1").Value = "Total Stock Volume"
    currentWS.Range("I1:L1").Font.Bold = True
    
    For i = 2 To finalRow
        currentTicker = currentWS.Cells(i, 1).Value
        nextTicker = currentWS.Cells(i + 1, 1).Value
        openPrice = currentWS.Cells(i, 3).Value
        closePrice = currentWS.Cells(i, 6).Value
        volume = currentWS.Cells(i, 7).Value
        
        volumeTotal = volumeTotal + volume
        
        If yearOpenPrice = -1 Then
            yearOpenPrice = openPrice
        End If
        
        If (nextTicker <> currentTicker) Then
            'Do calculations for price change and % percent change
            yearClosePrice = closePrice
            yearChange = yearClosePrice - yearOpenPrice
            percentChange = yearChange / yearOpenPrice
            
            'See if to store the min and max vlaues for the year
            If percentChange > maxIncrease Then
                maxIncrease = percentChange
                maxIncreaseTicker = currentTicker
            End If
            
            If percentChange < maxDecrease Then
                maxDecrease = percentChange
                maxDecreaseTicker = currentTicker
            End If
            
            If volumeTotal > maxVolume Then
                maxVolume = volumeTotal
                maxVolumeTicker = currentTicker
            End If
            
            Debug.Print ("Current Ticker/Total Volume " & currentTicker & " / " & volumeTotal)
            
            'Add values to spread sheet and format correct
            currentWS.Range("I" & j).Value = currentTicker
            currentWS.Range("J" & j).Value = yearChange
            currentWS.Range("K" & j).Value = percentChange
            currentWS.Range("K" & j).NumberFormat = "0.00%"
            currentWS.Range("L" & j).Value = volumeTotal
            
            'Change cell background based on if year change is positve or negative
            If yearChange < 0 Then
                currentWS.Range("J" & j).Interior.ColorIndex = myRed
            Else
                currentWS.Range("J" & j).Interior.ColorIndex = myGreen
            End If
            
            Debug.Print ("Processed Ticker: " & currentTicker)
            
            'Reset these variables
            yearOpenPrice = -1
            volumeTotal = 0
            
            j = j + 1
        End If
    Next i
    
    'Display the stocks with the max increase and decrease
    Debug.Print ("")
    Debug.Print ("Max Inc: " & maxIncreaseTicker & " / " & maxIncrease)
    Debug.Print ("Max Dec: " & maxDecreaseTicker & " / " & maxDecrease)
    
    'Running on Mac this line causes an overflow error
    'https://stackoverflow.com/questions/53809168/why-does-the-use-of-debug-print-lead-to-overflow-error
    Debug.Print ("Max Vol: " & maxVolumeTicker & " / " & maxVolume)
    
    'Display the header
    currentWS.Range("P1").Value = "Ticker"
    currentWS.Range("Q1").Value = "Value"
    currentWS.Range("P1,Q1").Font.Bold = True
    
    'Display values
    currentWS.Range("O2").Value = "Greatest % Increase"
    currentWS.Range("P2").Value = maxIncreaseTicker
    currentWS.Range("Q2").Value = maxIncrease
    
    currentWS.Range("O3").Value = "Greatest % Decrease"
    currentWS.Range("P3").Value = maxDecreaseTicker
    currentWS.Range("Q3").Value = maxDecrease
    
    currentWS.Range("O4").Value = "Greatest Total Volume"
    currentWS.Range("P4").Value = maxVolumeTicker
    currentWS.Range("Q4").Value = maxVolume
    
    'Format cells
    currentWS.Range("O2:O4").Font.Bold = True
    currentWS.Range("Q2,Q3").NumberFormat = "0.00%"
    
End Sub

