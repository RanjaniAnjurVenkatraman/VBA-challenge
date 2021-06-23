VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



Sub Stock():

    
    For Each ws In Worksheets

      
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

      
        Dim Ticker As String
        Dim Total_Volume As Double
        Total_Volume = 0
        Dim SummaryTable As Double
        SummaryTable = 2
        Dim Yearly_Change As Double
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim SummaryPrice As Double
        SummaryPrice = 2
        Dim Percent_Change As Double
        
        
        Dim LastRow As Double
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & SummaryTable).Value = Ticker
                ws.Range("L" & SummaryTable).Value = Total_Volume
                Total_Volume = 0
                
                Open_Price = ws.Range("C" & SummaryPrice)
                Close_Price = ws.Range("F" & i)
                Yearly_Change = Close_Price - Open_Price
                ws.Range("J" & SummaryTable).Value = Yearly_Change

               
                If Open_Price = 0 Then
                    Percent_Change = 0
                Else
                    Open_Price = ws.Range("C" & SummaryPrice)
                    Percent_Change = Yearly_Change / Open_Price
                End If
                
                ws.Range("K" & SummaryTable).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTable).Value = Percent_Change

                
                If ws.Range("J" & SummaryTable).Value >= 0 Then
                    ws.Range("J" & SummaryTable).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTable).Interior.ColorIndex = 3
                End If
                
                SummaryTable = SummaryTable + 1
                SummaryPrice = i + 1
            End If
        Next i
        
        Dim LastRow2 As Double
        LastRow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
            
        For i = 2 To LastRow2
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
            End If

            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value
            End If

            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
            End If

        Next i
       
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
       
        ws.Columns("I:Q").AutoFit

    Next ws

End Sub

