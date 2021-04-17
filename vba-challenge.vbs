Sub vba_challenge():

'Dimensions

    Dim Col  As Double
    Dim Total_Volume As Double
 
 
 
' Column Header


    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"

    Col = 2
    Cells(Col, 9).Value = Cells(Col, 1).Value

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For Row = 2 To LastRow

    If Cells(Row, 1).Value = Cells(Col, 9) Then
    
     
     Total_Volume = Total_Volume + Cells(Row, 7).Value
     
     Else
     
     Cells(Col, 10).Value = Total_Volume
     Total_Volume = Cells(Row, 7).Value
     Col = Col + 1
     Cells(Col, 9).Value = Cells(Row, 1).Value
     End If
     
     Next Row
     
     Cells(Col, 10).Value = Total_Volume
     
'Loop ws

    
    ' --------------------------------------------
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        ' Last Row
    
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Add Heading for summary
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        
        
        
        ' set dimensions
        
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Row = 2
        Dim column As Integer
        column = 1
        Dim i As Long
        
        
        
        ' Open Price
        
        Open_Price = Cells(2, column + 2).Value
        
        
         
         ' Loop
        
        For i = 2 To LastRow
       
            If Cells(i + 1, column).Value <> Cells(i, column).Value Then
              
                Ticker_Name = Cells(i, column).Value
                Cells(Row, column + 8).Value = Ticker_Name
             
                Close_Price = Cells(i, column + 5).Value
                '
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, column + 9).Value = Yearly_Change
            
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, column + 10).Value = Percent_Change
                    Cells(Row, column + 10).NumberFormat = "0.00%"
                End If
                
                Volume = Volume + Cells(i, column + 6).Value
                Cells(Row, column + 11).Value = Volume
                Row = Row + 1
                Open_Price = Cells(i + 1, column + 2)
                Volume = 0
            Else
                Volume = Volume + Cells(i, column + 6).Value
            End If
        Next i
        
        YCLastRow = WS.Cells(Rows.Count, column + 8).End(xlUp).Row
        
        
        ' Column Color
        
        For j = 2 To YCLastRow
            If (Cells(j, column + 9).Value > 0 Or Cells(j, column + 9).Value = 0) Then
                Cells(j, column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, column + 9).Value < 0 Then
                Cells(j, column + 9).Interior.ColorIndex = 3
            End If
        Next j
       
    
        
    Next WS
        
End Sub
