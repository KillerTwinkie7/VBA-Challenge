Sub sort_and_get_data():
    Dim firstOpen, lastClose As Double 'stores first and last values of ticker
    Dim volume() As Variant 'stores volume values in a list
    Dim x, a As Integer 'counter
    Dim length As Double 'stores length of column
    Dim totalVolume As Double 'stores the sum of all values in the volume column
    Dim ws As Worksheet 'defines worksheets as, well, worksheets
    
    
    For Each ws In ThisWorkbook.Worksheets 'Loop through each sheet in the workbook
        ws.Activate
        
        Cells(2, 18).Value = "Greatest % Increase"      '|
        Cells(3, 18).Value = "Greatest % Decrease"      '| Creates a small table to
        Cells(4, 18).Value = "Greatest Total Volume"    '| display superlative values
        Cells(1, 19).Value = "Ticker"                   '|
        Cells(1, 20).Value = "Value"                    '|
    
        ActiveSheet.Range("A:A").Copy                           '| Copies first column, then
        Range("I:I").Insert                                     '| puts them in column I while
        Range("I:I").RemoveDuplicates Columns:=1, Header:=xlYes '| removing duplicates.
        Cells(1, 9).Value = "Filtered tickers"                  '|
    
        Cells(1, 9).Value = "ticker"                '|
        Cells(1, 10).Value = "Opening Price"        '| Display Headers of Columns
        Cells(1, 11).Value = "Closing Price"        '|
        Cells(1, 12).Value = "Yearly Change"        '|
        Cells(1, 13).Value = "Percent Change"       '|
        Cells(1, 14).Value = "Total Stock Volume"   '|
    
        length = Cells(Rows.Count, "A").End(xlUp).Row 'gets length of first column
        x = 2 'counter initialization for display
        a = 1 'counter initialization for volume array
        Cells(2, 10).Value = Cells(2, 3).Value 'place first opening value in this cell
    
        For i = 2 To length 'loop through all of the data
            ReDim Preserve volume(0 To a) 'initialize the length of the dynamic list
            volume(a) = Cells(i, 7).Value 'append the newest value to the end of the list
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then 'when the values in the first column change, do the following
                lastClose = Cells(i, 6).Value 'get the final closing value
                firstOpen = Cells(i + 1, 3) 'get the first opening value of the next ticker
                totalVolume = Application.WorksheetFunction.Sum(volume) 'add all of the values of the volume list to get total
                Cells(x, 11).Value = lastClose 'display the last closing value to this cell
                Cells(x + 1, 10).Value = firstOpen 'display the next ticker's opening value to this cell
                Cells(x, 14).Value = totalVolume 'display total volume to this cell
                x = x + 1
            
                ReDim volume(0 To a) 'Clear list
                a = 0 'Reinitialize counter for volume storage
            End If
            a = a + 1
        Next i
        
        Range("N:N").NumberFormat = "#,###" 'formats total stock volume column to display commas
        
        length = Cells(Rows.Count, "I").End(xlUp).Row 'now get length of display table
        For i = 2 To length 'display the calculated yearly change and percent change values
            Cells(i, 12).Value = Cells(i, 11).Value - Cells(i, 10).Value 'calculate yearly change
            If Cells(i, 12).Value > 0 Then              '|
                Cells(i, 12).Interior.ColorIndex = "4"  '| Gives cells colors based on positive or negative
            ElseIf Cells(i, 12).Value < 0 Then          '| percent changes.
                Cells(i, 12).Interior.ColorIndex = "3"  '|
            End If
            Cells(i, 13).Value = (Cells(i, 11).Value / Cells(i, 10).Value) - 1 'calculate percent change
        Next i
        
        Range("M:M").NumberFormat = "0.00%" 'formats the precent change column
        
        Cells(2, 21).Value = Cells(2, 13).Value '| Set these cell values as the first values
        Cells(3, 21).Value = Cells(2, 13).Value '| to compare to future values and overwrite
        Cells(4, 21).Value = Cells(2, 14).Value '| them if certain conditions are met.
        
        For i = 2 To length
            If Cells(i, 13).Value > Cells(2, 21).Value Then '| Loop through the Percent Change column and
                Cells(2, 20).Value = Cells(i, 9)            '| find the largest. Put that value and the
                Cells(2, 21).Value = Cells(i, 13)           '| associated ticker value in the display table.
                Cells(2, 21).NumberFormat = "0.00%"         '|
            End If
            If Cells(i, 13).Value < Cells(3, 21).Value Then '| Loop through the Percent Change column and
                Cells(3, 20).Value = Cells(i, 9)            '| find the smallest. Put that value and the
                Cells(3, 21).Value = Cells(i, 13)           '| associated ticker value in the display table.
                Cells(3, 21).NumberFormat = "0.00%"         '|
            End If
            If Cells(i, 14).Value > Cells(4, 21) Then       '| Loop through the Total Stock Volume column
                Cells(4, 20).Value = Cells(i, 9)            '| and find the largest value. Put that value and
                Cells(4, 21).Value = Cells(i, 14)           '| the associated ticker value in the display table.
                Cells(4, 21).NumberFormat = "#,###"         '|
            End If
        Next i
        ActiveSheet.UsedRange.EntireColumn.AutoFit  '| Automatically changes column and row
        ActiveSheet.UsedRange.EntireRow.AutoFit     '| length to best fit data
    Next ws
End Sub