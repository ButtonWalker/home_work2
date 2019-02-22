Attribute VB_Name = "Module1"
Sub HM_WallStreet()

    'Declare the working variables
    Dim ws As Worksheet
    Dim Tkr As String
    Dim Stk_Vol As Double
    Dim Summary_Table_Row As Integer
    Dim open_Price As Double
    Dim close_Price As Double
    Dim yearly_Change As Double
    Dim percent_Change As Double

    'locate the open price cell
    open_Price = Cells(2, 3).Value
  
    For Each ws In Worksheets
    ws.Activate
    
    'Reset counter
    Stk_Vol = 0

    'Keep track of the location for each row/line in the summary table
    Summary_Table_Row = 2
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Find the Last Row in the worksheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To LastRow

            'check if we are still within the same value, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                'Set the Ticker
                Tkr = Cells(i, 1).Value

                'Add to Stk_Vol
                Stk_Vol = Stk_Vol + Cells(i, 7).Value
              
                'Print Ticker in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Tkr

                'Gather the closing price
                close_Price = Cells(i, 6).Value

                'yearly change calculation
                yearly_Change = (close_Price - open_Price)

                'Print the Yearly Change to the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = yearly_Change

                'Verify not dividing by 0
                If (open_Price = 0) Then

                    percent_Change = 0

                Else
                    
                    percent_Change = yearly_Change / open_Price
                
                End If

                 'Print the Percent Change to the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = percent_Change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

                'Print the Stk_Vol to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Stk_Vol

                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset Stk_Vol
                Stk_Vol = 0

            Else

                'Add to the Total Stock Volume
                Stk_Vol = Stk_Vol + Cells(i, 7).Value

            End If
              
        Next i

         lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
        'Color code yearly change
        
        For i = 2 To lastrow_summary_table
                If Cells(i, 10).Value > 0 Then
                    Cells(i, 10).Interior.ColorIndex = 10
                Else
                    Cells(i, 10).Interior.ColorIndex = 3
                End If
        Next i

    Next ws

End Sub

