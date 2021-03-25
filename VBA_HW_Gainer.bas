Attribute VB_Name = "Module1"

Sub tickersymbol():

'declaring worksheet count
Dim WS_Count As Integer
Dim i As Integer
Dim j As Long
Dim k As Long

'declaring lastrow worksheet count
Dim lastrowws As Long

'set variables for summary
Dim tickersymbol As String
Dim yearlychange As Double
Dim percyearlychange As Double
Dim Volume As LongLong

'summary table
Dim Summary_Table_Row As Long
Summary_Table_Row = 2

'set firstrow and lastrow as variable
Dim firstrow As Long
Dim lastrow As Long
firstrow = 2
lastrow = 2

'counting number of worksheets
WS_Count = ActiveWorkbook.Worksheets.Count

'for loop for worksheets
For i = 1 To WS_Count

'MsgBox (ActiveWorkbook.Worksheets(i).Name)


    'counting number of rows
    lastrowws = (ActiveWorkbook.Worksheets(i).Cells(Rows.Count, 1).End(xlUp).Row) + 1
    
    'putting headers
    Worksheets(i).Cells(1, 10).Value = "Ticker Symbol"
    Worksheets(i).Cells(1, 11).Value = "Yearly Change"
    Worksheets(i).Cells(1, 12).Value = "Percentage Change Yearly"
    Worksheets(i).Cells(1, 13).Value = "Total Stock Volume"
    Worksheets(i).Cells(2, 16).Value = "Greatest % increase"
    Worksheets(i).Cells(3, 16).Value = "Greatest % decrease"
    Worksheets(i).Cells(4, 16).Value = "Greatest total volume"
    Worksheets(i).Cells(1, 17).Value = "Ticker"
    Worksheets(i).Cells(1, 18).Value = "Value"
    
    
    'MsgBox (lastrow)


        For j = 2 To lastrowws
           
            If Worksheets(i).Cells(j + 1, 1).Value = Worksheets(i).Cells(j, 1).Value Then
                
                'add up volume and print
                Volume = Volume + Worksheets(i).Cells(j, 7).Value
                Worksheets(i).Range("M" & Summary_Table_Row).Value = Volume

            'Check if we are still within the same ticker symbol if not then set last row as the previous j
            ElseIf Worksheets(i).Cells(j + 1, 1).Value <> Worksheets(i).Cells(j, 1).Value Then
        
                'lastrow = j - 1
    
                'Set ticker symbol and print
                tickersymbol = Worksheets(i).Cells(j, 1).Value
                Worksheets(i).Range("J" & Summary_Table_Row).Value = tickersymbol
            
                'yearlychange and print
                yearlychange = Worksheets(i).Range("F" & j).Value - Worksheets(i).Range("C" & firstrow).Value
                Worksheets(i).Range("K" & Summary_Table_Row).Value = yearlychange
                If Worksheets(i).Range("K" & Summary_Table_Row).Value > 0 Then
                Worksheets(i).Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                Worksheets(i).Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                Volume = Volume + Worksheets(i).Cells(j, 7).Value
                Worksheets(i).Range("M" & Summary_Table_Row).Value = Volume
                
            
                'percentagechangeyearly and print
                If Worksheets(i).Cells(firstrow, 3).Value = 0 Then
                percyyearlychange = 1
                Worksheets(i).Range("L" & Summary_Table_Row).Value = percyearlychange
                Worksheets(i).Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
              
                
                Else
                percyearlychange = yearlychange / Worksheets(i).Cells(firstrow, 3).Value
                Worksheets(i).Range("L" & Summary_Table_Row).Value = percyearlychange
                Worksheets(i).Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                
                firstrow = j + 1
                
                End If
                
            If Worksheets(i).Cells(Summary_Table_Row, 12).Value > Worksheets(i).Cells(2, 18).Value Then
            
                Worksheets(i).Cells(2, 18).Value = Worksheets(i).Cells(Summary_Table_Row, 12).Value
                Worksheets(i).Cells(2, 17).Value = Worksheets(i).Cells(Summary_Table_Row, 10).Value
                Worksheets(i).Cells(2, 18).NumberFormat = "0.00%"
                
                Else
        
            End If
               
            If Worksheets(i).Cells(Summary_Table_Row, 12).Value < Worksheets(i).Cells(3, 18).Value Then
            
                Worksheets(i).Cells(3, 18).Value = Worksheets(i).Cells(Summary_Table_Row, 12).Value
                Worksheets(i).Cells(3, 17).Value = Worksheets(i).Cells(Summary_Table_Row, 10).Value
                Worksheets(i).Cells(3, 18).NumberFormat = "0.00%"
                
                Else
        
            End If
                
                
            If Worksheets(i).Cells(Summary_Table_Row, 13).Value > Worksheets(i).Cells(4, 18).Value Then
            
                Worksheets(i).Cells(4, 18).Value = Worksheets(i).Cells(Summary_Table_Row, 13).Value
                Worksheets(i).Cells(4, 17).Value = Worksheets(i).Cells(Summary_Table_Row, 10).Value
                
                Else
        
            End If
        'add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        Volume = 0


        End If
    
        Next j
        
        Summary_Table_Row = 2
        firstrow = 2
        
         
Next i

End Sub


