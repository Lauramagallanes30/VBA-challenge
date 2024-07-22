Attribute VB_Name = "Module1"
Sub stock_data()
    Dim i As Long
    Dim RowCount As Long
    Dim OpenPrice As Double
    Dim ClosingPrice As Double
    Dim TotalVol As Double
    Dim j As Long
    Dim Quarterly As Double
    Dim ws As Worksheet
    Dim PerChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim condition1 As FormatCondition
    Dim condition2 As FormatCondition

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ' Add Summary Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Add headers for calculated values
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest Increase"
        ws.Cells(3, 14).Value = "Greatest Decrease"
        ws.Cells(4, 14).Value = "Greatest Volume"
        
        ' Find the last row of data
        RowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Initialize variables
        j = 2
        TotalVol = 0
        greatestIncrease = -999999 ' Initialize with a very low value
        greatestDecrease = 999999  ' Initialize with a very high negative value
        greatestVolume = 0
        
        OpenPrice = ws.Cells(2, 3).Value

        For i = 2 To RowCount
            ' Check if the cell value can be converted to a Double before adding to TotalVol
            If IsNumeric(ws.Cells(i, 7).Value) Then
                TotalVol = TotalVol + CDbl(ws.Cells(i, 7).Value)
            End If
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ClosingPrice = ws.Cells(i, 6).Value
                Quarterly = ClosingPrice - OpenPrice
                PerChange = (Quarterly / OpenPrice) * 100

                ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(j, 10).Value = Quarterly
                ws.Cells(j, 11).Value = Format(PerChange, "0.00") & "%"
                ws.Cells(j, 12).Value = TotalVol

                ' Determine greatest increase, decrease, and volume
                If PerChange > greatestIncrease Then
                    greatestIncrease = PerChange
                    greatestIncreaseTicker = ws.Cells(i, 1).Value
                End If

                If PerChange < greatestDecrease Then
                    greatestDecrease = PerChange
                    greatestDecreaseTicker = ws.Cells(i, 1).Value
                End If

                If TotalVol > greatestVolume Then
                    greatestVolume = TotalVol
                    greatestVolumeTicker = ws.Cells(i, 1).Value
                End If

                j = j + 1
                OpenPrice = ws.Cells(i + 1, 3).Value
                TotalVol = 0
            End If
        Next i
        
        ' Output greatest values
        ws.Cells(2, 15).Value = greatestIncreaseTicker
        ws.Cells(2, 16).Value = Format(greatestIncrease, "0.00") & "%"

        ws.Cells(3, 15).Value = greatestDecreaseTicker
        ws.Cells(3, 16).Value = Format(greatestDecrease, "0.00") & "%"

        ws.Cells(4, 15).Value = greatestVolumeTicker
        ws.Cells(4, 16).Value = greatestVolume

        ' Apply conditional formatting
        Set condition1 = ws.Range("J2:J" & j - 1).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        Set condition1 = ws.Range("J2:J" & j - 1).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        condition1.Interior.Color = RGB(0, 255, 0)

        Set condition2 = ws.Range("J2:J" & j - 1).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
        condition2.Interior.Color = RGB(255, 0, 0)

        Set condition1 = ws.Range("K2:K" & j - 1).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        condition1.Interior.Color = RGB(0, 255, 0)

        Set condition2 = ws.Range("K2:K" & j - 1).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
        condition2.Interior.Color = RGB(255, 0, 0)
    Next ws
End Sub
