Option Explicit

Sub TableWallStreet()
Dim tbl As ListObject
Dim dtbl As ListObject 'Table variable to allow easier reference
Dim stbl As ListObject
Dim gtbl As ListObject
Dim DataTableRange As Range 'store the  range to be converted into data table
Dim DataTableName As String 'store data table name on each sheet
Dim SumTableRange As Range
Dim SumTableName As String
Dim GreTableRange As Range
Dim GreTableName As String
Dim i As Integer 'worksheet counter
Dim c As Integer 'last column of data table to offset summary table
Dim sc As Integer 'last column of summary data table to offset Greatest Table
Dim SumHead As String 'Store summary table heading
Dim GreHead As String
Dim x As Integer 'Sum table header locations
Dim y As Integer 'Great table header locations
Dim r As Double 'Row counter for data table loop
Dim Ticker As Range 'Ticker current row
Dim TickerA As Range 'Ticker row above
Dim TickerB As Range 'Ticker row below
Dim Volume As Range 'Ticker Volume location
Dim OpenP As Range 'Ticker Opeing price location
Dim CloseP As Range 'Ticker Closing price location
Dim opn As Currency 'Opening ticker price
Dim cls As Currency 'Closing ticker price
Dim tot_vol As Double 'Volume counter
Dim sr As Integer 'Summary table row
Dim TS As Range ' Summary table header
Dim YC As Range 'Summary table header
Dim PC As Range 'Summary table header
Dim TSV As Range 'Summary table header
Dim ws As Worksheet
Dim maxinc As Double
Dim maxincR As String
Dim maxdec As Double
Dim maxdecR As String
Dim maxvol As Double
Dim maxvolR As String

'Speed up code
Application.ScreenUpdating = False

    'create loop counter up to the number of worksheets
    For i = 1 To ActiveWorkbook.Worksheets.Count
    
        'start the loop by opening the sheet in the loop
        ActiveWorkbook.Worksheets(i).Activate
        
        Set ws = Worksheets(i)
        
              
        'name the upcoming table based on the selected sheet name
        DataTableName = ws.Name & "Data"
        
        'Set the range that will turned into the table (current region)
        Set DataTableRange = ws.Range("A1").CurrentRegion
        
        'convert the range into a table using the variables above
        Sheets(i).ListObjects.Add(SourceType:=xlSrcRange, _
        Source:=DataTableRange, _
        xlListObjectHasHeaders:=xlYes _
        ).Name = DataTableName
        
        'Store the Data Table as a variable - for easier coding
        Set dtbl = Sheets(i).ListObjects(DataTableName)
        
        'Create summary table with one column gap between data table
        c = ws.ListObjects(DataTableName).HeaderRowRange.Count
            
            'Enter summary table headers
            For x = 2 To 5
                                    
            Select Case x
                Case 2
                    SumHead = "Ticker Symbol"
                Case 3
                    SumHead = "Yearly Change"
                Case 4
                    SumHead = "Percent Change"
                Case 5
                    SumHead = "Total Stock Volume"
            End Select
            
            Sheets(i).ListObjects(DataTableName).HeaderRowRange(c).Offset(, x).Value = SumHead
                
            Next x
            
        'convert headers into summary table
        Set SumTableRange = ws.ListObjects(DataTableName).HeaderRowRange(c).Offset(, 2).CurrentRegion
        
        SumTableName = ws.Name & "Summary"
        
        'convert the range into a table using the variables above
        Sheets(i).ListObjects.Add(SourceType:=xlSrcRange, _
        Source:=SumTableRange, _
        xlListObjectHasHeaders:=xlYes _
        ).Name = SumTableName
        
        'store summary table into variable for easier coding later
        Set stbl = ws.ListObjects(SumTableName)
                
        'set variable values in prep for looping through data table
        tot_vol = 0
        sr = 1
        
        'loop through data table rows to populate summary table
        For r = 1 To dtbl.ListRows.Count
        
        'Set the variables to refer to parts of the data table to easier read and understand coding during conditionals
        Set Ticker = dtbl.DataBodyRange.Cells(r, dtbl.ListColumns("<ticker>").Index)
        Set TickerA = dtbl.DataBodyRange.Cells(r - 1, dtbl.ListColumns("<ticker>").Index)
        Set TickerB = dtbl.DataBodyRange.Cells(r + 1, dtbl.ListColumns("<ticker>").Index)
        Set Volume = dtbl.DataBodyRange.Cells(r, dtbl.ListColumns("<vol>").Index)
        Set OpenP = dtbl.DataBodyRange.Cells(r, dtbl.ListColumns("<open>").Index)
        Set CloseP = dtbl.DataBodyRange.Cells(r, dtbl.ListColumns("<close>").Index)
        
        'If the ticker name is the same as the name below and above, then it just adds to the volume for this ticker name
        If Ticker.Value = TickerB.Value And Ticker.Value = TickerA.Value Then
            tot_vol = tot_vol + Volume.Value
            
            
        'if the ticker name is same as the ticker below but different to the one above then _
        it is the opening price for that ticker
        ElseIf Ticker.Value = TickerB.Value And Ticker.Value <> TickerA.Value Then
        
        tot_vol = tot_vol + Volume.Value 'start adding to the volume tally
        opn = OpenP.Value 'store the opening price in opn variable
        
        'if the ticker name is the same as the ticker above but different to the name below then _
        it is the closing price for that ticker
        ElseIf Ticker.Value = TickerA.Value And Ticker.Value <> TickerB.Value Then
        
        tot_vol = tot_vol + Volume.Value 'last value to be added to the volume tally
        cls = CloseP.Value 'store the closing price in cls variable
        
        'summary table:
        stbl.ListRows.Add (sr)
        Set TS = stbl.DataBodyRange.Cells(sr, stbl.ListColumns("Ticker Symbol").Index)
        Set YC = stbl.DataBodyRange.Cells(sr, stbl.ListColumns("Yearly Change").Index)
        Set PC = stbl.DataBodyRange.Cells(sr, stbl.ListColumns("Percent Change").Index)
        Set TSV = stbl.DataBodyRange.Cells(sr, stbl.ListColumns("Total Stock Volume").Index)
       
        
        TS.Value = Ticker.Value 'Insert the ticker name to summary table
        YC.Value = opn - cls 'insert the difference bw open and close price to summary table
        If opn = 0 And cls = 0 Then 'Catch in case opn is 0
        PC.Value = 0
        Else
        PC.Value = (opn - cls) / opn 'insert % change to summary table
        End If
        If tot_vol = 0 Then
        TSV.Value = 0
        Else
        TSV.Value = tot_vol 'insert total volume tally into summary table
        End If
        tot_vol = 0 'reset volumn tally
        sr = sr + 1 'add to summary table row counter for next time
        
        
        'no need to reset opn or cls will be redefined during if statement later
               
        End If
        
        Next r
    
    'Formatting of summary table
    stbl.ListColumns("Percent Change").DataBodyRange.NumberFormat = "0.00%"
    stbl.ListColumns("Total Stock Volume").DataBodyRange.NumberFormat = "0,00"
    stbl.ListColumns("Yearly Change").DataBodyRange.FormatConditions.Add xlCellValue, xlGreater, 0
    stbl.ListColumns("Yearly Change").DataBodyRange.FormatConditions(1).Interior.Color = vbGreen
    stbl.ListColumns("Yearly Change").DataBodyRange.FormatConditions.Add xlCellValue, xlLess, 0
    stbl.ListColumns("Yearly Change").DataBodyRange.FormatConditions(2).Interior.Color = vbRed
    
    'Create Greatest Table
    sc = stbl.ListColumns.Count
    'Enter Greatest table headers
            For y = 2 To 4
                                    
            Select Case y
                Case 2
                    GreHead = "Criteria"
                Case 3
                    GreHead = "Ticker"
                Case 4
                    GreHead = "Value"
                
            End Select
            
            stbl.HeaderRowRange(sc).Offset(, y).Value = GreHead
                
            Next y
    'Create Greatest Table
    Set GreTableRange = stbl.HeaderRowRange(sc).Offset(, 2).CurrentRegion
    GreTableName = ws.Name & "Greatest"
    
    ws.ListObjects.Add(SourceType:=xlSrcRange, _
        Source:=GreTableRange, _
        xlListObjectHasHeaders:=xlYes _
        ).Name = GreTableName
    
    Set gtbl = ws.ListObjects(GreTableName)
   For r = 1 To 3
    gtbl.ListRows.Add (r)
    Select Case r
        Case Is = 1
    gtbl.DataBodyRange.Cells(r, 1).Value = "Greatest % Increase"
        Case Is = 2
    gtbl.DataBodyRange.Cells(r, 1).Value = "Greatest % Decrease"
        Case Is = 3
    gtbl.DataBodyRange.Cells(r, 1).Value = "Greatest Total Volume"
    End Select
   Next r
    
    maxinc = 0
    For r = 1 To stbl.ListRows.Count
        If stbl.DataBodyRange.Cells(r, stbl.ListColumns("Percent Change").Index).Value > maxinc Then
        
        maxinc = stbl.DataBodyRange.Cells(r, stbl.ListColumns("Percent Change").Index).Value
        maxincR = stbl.DataBodyRange.Cells(r, stbl.ListColumns("Ticker Symbol").Index).Value
        
        End If
        
    Next r
     
    maxdec = 0
    For r = 1 To stbl.ListRows.Count
        If stbl.DataBodyRange.Cells(r, stbl.ListColumns("Percent Change").Index).Value < maxdec Then
        
        maxdec = stbl.DataBodyRange.Cells(r, stbl.ListColumns("Percent Change").Index).Value
        maxdecR = stbl.DataBodyRange.Cells(r, stbl.ListColumns("Ticker Symbol").Index).Value
        
        End If
        
    Next r
      
    maxvol = 0
    For r = 1 To stbl.ListRows.Count
        If stbl.DataBodyRange.Cells(r, stbl.ListColumns("Total Stock Volume").Index).Value > maxvol Then
        
        maxvol = stbl.DataBodyRange.Cells(r, stbl.ListColumns("Total Stock Volume").Index).Value
        maxvolR = stbl.DataBodyRange.Cells(r, stbl.ListColumns("Ticker Symbol").Index).Value
        End If
        
    Next r
      
    '% increase ticker
    gtbl.DataBodyRange.Cells(1, 2).Value = maxincR
    '% decrease ticker
    gtbl.DataBodyRange.Cells(2, 2).Value = maxdecR
    'max volume ticker
    gtbl.DataBodyRange.Cells(3, 2).Value = maxvolR
    
    '%increase
    gtbl.DataBodyRange.Cells(1, 3).Value = maxinc
    '%decrease
    gtbl.DataBodyRange.Cells(2, 3).Value = maxdec
    'max volume
    gtbl.DataBodyRange.Cells(3, 3).Value = maxvol
    
    gtbl.DataBodyRange.Cells(1, 3).NumberFormat = "0.00%"
    gtbl.DataBodyRange.Cells(2, 3).NumberFormat = "0.00%"
    gtbl.DataBodyRange.Cells(3, 3).NumberFormat = "0,00"
    gtbl.Range.EntireColumn.AutoFit
    
    Next i

Application.ScreenUpdating = True

End Sub
