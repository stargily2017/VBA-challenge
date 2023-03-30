# VBA-challenge

Sub StockMarket1()


'Define a variable for year open

Dim open_price As Double

'Define avariable for year close

Dim close_price As Double

'Define a variable for yearly change

Dim Yearly_Change As Double

'Define a variable for total stock volume

Dim Total_Stock_Volume As Double

'Define a variable for percent change

Dim Percent_Change As Double

'Define a variable to set up a row to start

Dim First_Value As Integer


'Define a variable for Ticker

Dim Ticker As String

'Define variable of the worksheet to excute the code in all work sheet at once in the workbook

Dim ws As Worksheet

For Each ws In Worksheets

    'Assign a column header for each names

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    'Assign intiger for the loop to start
    First_Value = 2
    Row_index = 1
    Total_Stock_Volume = 0

    'Go to the last row of coumn A

    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        'For each Ticker summrize and loop the yearly change, percent change, and total stock volume

        For i = 2 To LastRow

            'If Tickersymbol change or not equal to the previous one excute to record

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Get the Tickersymbol

            Ticker = ws.Cells(i, 1).Value

            'Intiate the variable to go to the next Ticker Alphabet

            Row_index = Row_index + 1

            ' Get the first value  from "C" and last close value from"F"

            open_price = ws.Cells(Row_index, 3).Value
            close_price = ws.Cells(i, 6).Value

            ' A for loop to sum the total stock volume using vol which is found in column 7 or "G"

            For j = Row_index To i

                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value

            Next j

            'When the loop get the value zero open the data

            If open_price = 0 Then

                Percent_Change = close_price

            Else
                Yearly_Change = close_price - open_price

                Percent_Change = Yearly_Change / open_price

            End If
         

            ' To Get the values in the worksheet

            ws.Cells(First_Value, 9).Value = Ticker
            ws.Cells(First_Value, 10).Value = Yearly_Change
            ws.Cells(First_Value, 11).Value = Percent_Change
            

            'Use percentage format

            ws.Cells(First_Value, 11).NumberFormat = "0.00%"
            ws.Cells(First_Value, 12).Value = Total_Stock_Volume

            'In the data summery when the first row task completed go to the next row

            First_Value = First_Value + 1

            'Get back the variable to zero

            Total_Stock_Volume = 0
            Yearly_Change = 0
            Percent_Change = 0

            'Move i number to variable previous_i
            Row_index = i

        End If

    'Done the loop

    Next i

'Get the lastrow in "j"

    LastRowj = ws.Cells(Rows.Count, "J").End(xlUp).Row


        For n = 2 To LastRowj
           
            If ws.Cells(n, 10) > 0 Then

                ws.Cells(n, 10).Interior.ColorIndex = 4

            Else

                ws.Cells(n, 10).Interior.ColorIndex = 3
            End If

        Next n




    LastRowk = ws.Cells(Rows.Count, "K").End(xlUp).Row

    'Define variable to initiate the second summery table value

    Increase = 0
    Decrease = 0
    Greatest = 0

        'find max/min for percentage change and the max volume Loop
        For p = 3 To LastRowk

            'increment
            
            End_v = p - 1

            'cuurrent value
            
            Actual_v = ws.Cells(p, 11).Value

           'previous percent
           
            previous_v = ws.Cells(End_v, 11).Value

            'greatest total volume row
            volume = ws.Cells(p, 12).Value

           
            previous_volume = ws.Cells(End_v, 12).Value


            'Find the increase gretest % increase
            
            If Increase > Actual_v And Increase > previous_v Then

                Increase = Increase

               
            ElseIf Actual_v > Increase And Actual_v > prevous_v Then

                Increase = Actual_v

                'define name for increase percentage
                increase_Ticker = ws.Cells(p, 9).Value

            ElseIf previous_v > Increase And previous_v > Actual_v Then

                Increase = prevous_v

                'define name for increase percentage
                increase_Ticker = ws.Cells(End_v, 9).Value

            End If

      'find Decrease percent

            If Decrease < Actual_v And Decrease < previous_v Then

                
                Decrease = Decrease


            ElseIf Actual_v < Increase And Actual_v < previous_v Then

                Decrease = Actual_v


                decrease_Ticker = ws.Cells(p, 9).Value

            ElseIf previous_v < Increase And previous_v < Actual_v Then

                Decrease = previous_v

                decrease_Ticker = ws.Cells(End_v, 9).Value

            End If

      'find the volume
      
            If Greatest > volume And Greatest > prevous_volume Then

                Greatest = Greatest


            ElseIf volume > Greatest And volume > prevous_volume Then

                Greatest = volume

                'greatest volume
                
                greatest_Ticker = ws.Cells(p, 9).Value

            ElseIf prevous_volume > Greatest And prevous_volume > volume Then

                Greatest = prevous_volume

                'define name for greatest volume
                greatest_Ticker = ws.Cells(End_v, 9).Value

            End If

        Next p
 
    'Get for greatest increase, greatest increase, and  greatest volume Ticker name
    ws.Range("P2").Value = increase_Ticker
    ws.Range("P3").Value = decrease_Ticker
    ws.Range("P4").Value = greatest_Ticker
    ws.Range("Q2").Value = Increase
    ws.Range("Q3").Value = Decrease
    ws.Range("Q4").Value = Greatest

    'Greatest increase and decrease in percentage format

    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"

'Execute to next worksheet
Next ws

End Sub
