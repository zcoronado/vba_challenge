' Create a script that will loop through all the stocks for one year and output the following information.
' The ticker symbol.
' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' The total stock volume of the stock.
' You should also have conditional formatting that will highlight positive change in green and negative change in red.

Sub Tickercalculation()

' Data
' Col 1 - Ticket
' Col 2 - Date
' Col 3 - Open
' Col 4 - High
' Col 5 - Low
' Col 6 - Close
' Col 7 - Volume
' Col 8 - Volume in 1000's


' Output
' Col 9 - Ticker
' Col 10 - Yearly Change
' Col 11 - Percent Change
' Col 12 - Total Stock Volume


Dim datarow As Long
Dim outputrow As Long
Dim sheetnum As Long



Dim openprice As Double
Dim totalstockvolume As Double
Dim closeprice As Double




For sheetnum = 1 To Worksheets.Count

    Worksheets(sheetnum).Activate
    outputrow = 2

    openprice = ActiveSheet.Range("C2").Value

    'Titles
    
    ActiveSheet.Cells(1, 9).Value = "Ticket"
    ActiveSheet.Cells(1, 10).Value = "Yearly Change"
    ActiveSheet.Cells(1, 11).Value = "Percent Change"
    ActiveSheet.Cells(1, 12).Value = "Total Stock Volume"

    ' Start loop at A2
    For datarow = 2 To ActiveSheet.Range("A2").End(xlDown).Row
        If ActiveSheet.Cells(datarow, 1).Value <> ActiveSheet.Cells(datarow + 1, 1).Value Then
            ' Now at the edge
            ' add what is in Col g to the total stock counter
            totalstockvolume = totalstockvolume + ActiveSheet.Cells(datarow, 8).Value
            ' grab the closing price from Col F
            closeprice = ActiveSheet.Cells(datarow, 6).Value

            ' percent change
            If openprice = 0 Then
                ActiveSheet.Cells(outputrow, 11).Value = "NaN"
            Else
                ActiveSheet.Cells(outputrow, 11).Value = (closeprice - openprice) / openprice
            End If

            'change to percent
            Columns("K:K").Select
            Selection.Style = "Percent"
            Selection.NumberFormat = "0.0%"
            Selection.NumberFormat = "0.00%"


            ' yearly change
            ActiveSheet.Cells(outputrow, 10).Value = closeprice - openprice



            ' total stock volume
            ActiveSheet.Cells(outputrow, 12).Value = totalstockvolume

            ' ticker
            ActiveSheet.Cells(outputrow, 9).Value = ActiveSheet.Cells(datarow, 1).Value

            ' calc yearly change as close_price - open_price
            ' calc percentage change as close_price - open_price / open_price
            ' check denominator is not 0
            ' Copy value from Col A to Col I
            ' Then dump the yearly change, percentage change,
            ' Add 1 to the row counter for the output table

            outputrow = outputrow + 1

            ' Updat the open price to be the open price of the new ticker
            totalstockvolume = 0
            openprice = ActiveSheet.Cells(datarow + 1, 3).Value


        Else
            ' if it is not at the edge then
            ' dont change the open value
            ' add whatever is in Col G to the total stock volume counter
            totalstockvolume = totalstockvolume + ActiveSheet.Cells(datarow, 8).Value

        End If

    Next datarow


Next sheetnum

End Sub