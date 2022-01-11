Attribute VB_Name = "Module4"

Sub YearStock():
Dim sht As Worksheet
'Set sht = ActiveSheet
Dim Brand_Name As String
Dim Brand_Summary_Row As Integer
Brand_Summary_Row = 2
Brand_Summary_Column = 1
'row_count = Cells(Rows.Count, "A").End(xlUp).Row
'LastRow = sht.Range("A1").CurrentRegion.Rows.Count
Dim LastRow As Long
Dim Open_Value As Double
Dim Close_Value As Double
Dim Yearly_Change As Double
Dim total_volume As Double
'Open_Value = Cells(2, "C").Value
tolal_volume = 0
'Set sht = ActiveSheet
'LastRow = sht.Cells.SpecialCells(xlCellTypeLastCell).Row


    For Each sht In Worksheets
        sht.Select
        'Set sht = ActiveSheet
        ''LastRow = sht.Cells.SpecialCells(xlCellTypeLastCell).Row
        LastRow = Cells.SpecialCells(xlCellTypeLastCell).Row
        Brand_Summary_Row = 2   ' FIX
        Open_Value = Cells(2, "C").Value
        
        For i = 2 To LastRow
        
        total_volume = Cells(i, "G").Value + total_volume
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Brand_Name = Cells(i, 1).Value
                Close_Value = Cells(i, "F").Value
                Yearly_Change = Close_Value - Open_Value
                Cells(Brand_Summary_Row, "J").Value = Yearly_Change
                    If Open_Value = 0 Then
                        Cells(Brand_Summary_Row, "K").Value = 0
                        Else
                        Cells(Brand_Summary_Row, "K").Value = FormatPercent(Yearly_Change / Open_Value, 2)
                            If Cells(Brand_Summary_Row, "J").Value < 0 Then
                                Cells(Brand_Summary_Row, "J").Interior.ColorIndex = 3
                                Else
                                Cells(Brand_Summary_Row, "J").Interior.ColorIndex = 4
                            End If
                    End If
                Cells(Brand_Summary_Row, "L").Value = total_volume
                Open_Value = Cells(1 + i, "C").Value
                total_volume = 0
                Cells(Brand_Summary_Row, "I").Value = Brand_Name
                'Range("I" & Brand_Summary_Row).Value = Brand_Name
                Brand_Summary_Row = Brand_Summary_Row + 1
            End If
        Next i
        LastRow = 0     ' FIX
    Next                ' FIX

    
End Sub

