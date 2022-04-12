# CRangeSum
PersonalFormular: 特定方式数据求和，用于批量关店模板
Option Base 1


Function CRange(ByRef myRange As Range, ByRef RowTitle As Variant, ByRef ColTitle As Variant, Optional NRC As Integer)
'满足 行名和列名 的条件范围
'NRC 1列标号为数字；2 行标号为数字；3 行标号和列标号都是数字; 11 整列； 22 整行； 33 整区

Dim Brr
Dim i, Ti As Long:  Ti = 0
Dim j, Tj As Long:  Tj = 0
Dim myC As Long: myC = 0
Dim myRow, myCol As Long
'    myRow = -1:    myCol = -1

Select Case NRC
    Case 1  '1列标号为数字
        myCol = ColTitle:
        Tj = 1
        For i = 1 To myRange.Rows.Count
            If myRange.Cells(i, 1).Value = RowTitle Then Ti = Ti + 1
        Next
        ReDim Brr(Ti * Tj)
        For i = 1 To myRange.Rows.Count
            If myRange.Cells(i, 1).Value = RowTitle Then
                myC = myC + 1
                Brr(myC) = IIf(IsNumeric(myRange.Cells(i, myCol)), myRange.Cells(i, myCol), 0)
            End If
        Next
    Case 11 '11 整列
        If RowTitle = 1 Then
            myCol = ColTitle
        Else
            For j = 1 To myRange.Columns.Count
                If myRange.Cells(1, j).Value = ColTitle Then myCol = j
            Next
        End If
        Tj = 1
        Ti = myRange.Rows.Count
        ReDim Brr(Ti * Tj)
        For i = 1 To Ti
            myC = myC + 1
            Brr(myC) = IIf(IsNumeric(myRange.Cells(i, myCol)), myRange.Cells(i, myCol), 0)
        Next
    Case 2  '2 行标号为数字
        myRow = RowTitle
        Ti = 1
        For j = 1 To myRange.Columns.Count
            If myRange.Cells(1, j).Value = ColTitle Then Tj = Tj + 1
        Next
        ReDim Brr(Ti * Tj)
        For j = 1 To myRange.Columns.Count
            If myRange.Cells(1, j).Value = ColTitle Then
                myC = myC + 1
                Brr(myC) = IIf(IsNumeric(myRange.Cells(myRow, j)), myRange.Cells(myRow, j), 0)
            End If
        Next
    Case 22  '22 整行
        If ColTitle = 2 Then
            myRow = RowTitle
        Else
            For i = 1 To myRange.Rows.Count
                If myRange.Cells(i, 1).Value = RowTitle Then myRow = i
            Next
        End If
        Ti = 1
        Tj = myRange.Columns.Count
        ReDim Brr(Ti * Tj)
        For j = 1 To Tj
            myC = myC + 1
            Brr(myC) = IIf(IsNumeric(myRange.Cells(myRow, j)), myRange.Cells(myRow, j), 0)
        Next
    Case 3  '3 行标号和列标号都是数字
        myCol = ColTitle
        myRow = RowTitle
        Brr(1) = myRange.Cells(myRow, myCol)
     Case 33    '33 整区
        Ti = myRange.Rows.Count
        Tj = myRange.Columns.Count
        ReDim Brr(Ti * Tj)
        For i = 1 To Ti
        For j = 1 To Tj
            myC = myC + 1
            Brr(myC) = IIf(IsNumeric(myRange.Cells(i, j)), myRange.Cells(i, j), 0)
        Next
        Next
   Case Else
        For i = 1 To myRange.Rows.Count
            If myRange.Cells(i, 1).Value = RowTitle Then Ti = Ti + 1
        Next
        For j = 1 To myRange.Columns.Count
            If myRange.Cells(1, j).Value = ColTitle Then Tj = Tj + 1
        Next
        ReDim Brr(Ti * Tj)

        For i = 1 To myRange.Rows.Count
            If myRange.Cells(i, 1).Value = RowTitle Then
            For j = 1 To myRange.Columns.Count
                If myRange.Cells(1, j).Value = ColTitle Then
                    myC = myC + 1
                    Brr(myC) = IIf(IsNumeric(myRange.Cells(i, j)), myRange.Cells(i, j), 0)
                End If
            Next
            End If
        Next
End Select

CRange = Brr

End Function
