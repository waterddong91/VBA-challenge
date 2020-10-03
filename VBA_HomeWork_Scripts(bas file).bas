Attribute VB_Name = "Module1"
Sub VBA_WallStreet():


    'Loop through all the stocks for one year
    For Each ws In Worksheets


        'Column Headers / Data Field Labels
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        


        'Set/Declare Variables
        Dim Ticker As String
        Dim LastRow As Long
        Dim Total_Volume As Double
        Total_Volume = 0
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim First_Price As Long
        First_Price = 2
        Dim Percent_Change As Double
        

        'Set/Delare the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        



        For i = 2 To LastRow


            'Set Total_Volume
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
           

            'Check the status
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


                'Set Ticker
                Ticker = ws.Cells(i, 1).Value
                'Print Ticker in the table
                ws.Range("I" & SummaryTableRow).Value = Ticker
                'Print Total_Volume in the table
                ws.Range("L" & SummaryTableRow).Value = Total_Volume
                'Reset Total_Volume
                Total_Volume = 0


                
                'Set Open/Close Price and Yealry Change
                Open_Price = ws.Range("C" & First_Price)
                Close_Price = ws.Range("F" & i)
                Yearly_Change = Close_Price - Open_Price
                ws.Range("J" & SummaryTableRow).Value = Yearly_Change


                'Percent Change
                If Open_Price = 0 Then
                    Percent_Change = 0
                
                Else
                    Open_Price = ws.Range("C" & First_Price)
                    Percent_Change = Yearly_Change / Open_Price
                
                End If
                
                'Put % Symbol with 2 Decimal points
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = Percent_Change


                'Put Highlight
                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                
                End If
                            
                ' Keep add one row to the table
                SummaryTableRow = SummaryTableRow + 1
                First_Price = i + 1
                
                End If
            
            Next i


' #CHALLENGES#

        ' Column Headers / Data Field Labels
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        
        'Set/Declare Variables
        Dim Greatest_Increase As Double
        Greatest_Increase = 0
        Dim Greatest_Decrease As Double
        Greatest_Decrease = 0
        Dim LastRow2 As Long
        Dim Greatest_Total_Volume As Double
        Greatest_Total_Volume = 0



            'Set Last low again
            LastRow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            
            'Loop
            For i = 2 To LastRow
                
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
        
        'Put % Symbol with 2 Decimal points
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"



    Next ws


End Sub

