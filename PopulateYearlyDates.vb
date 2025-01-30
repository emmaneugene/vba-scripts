Sub PopulateYearlyDates()
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    Dim startDate As Date
    Dim currentDate As Date
    Dim lastRow As Long
    Dim currentRow As Long
    Dim yearMergeStartRow As Long
    Dim monthMergeStartRow As Long
    Dim previousYear As String
    Dim previousMonth As String

    ' Set the working sheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Get the year input from user
    Dim yearInput As String
    yearInput = InputBox("Enter the year (YYYY):", "Year Input")

    ' Validate year input
    If Not IsNumeric(yearInput) Or Len(yearInput) <> 4 Then
        MsgBox "Please enter a valid year in YYYY format.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Set start date to January 1st of input year
    startDate = DateSerial(CInt(yearInput), 1, 1)

     lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 1 Then
        lastRow = 0
    Else
        ' If the last cell is part of a merged range, get the last cell of the merge
        If ws.Cells(lastRow, "A").MergeCells Then
            Dim mergedRange As Range
            Set mergedRange = ws.Cells(lastRow, "A").MergeArea
            lastRow = mergedRange.Row + mergedRange.Rows.Count - 1
        End If

        ' Do the same check for column B (months)
        Dim lastRowB As Long
        lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        If ws.Cells(lastRowB, "B").MergeCells Then
            Set mergedRange = ws.Cells(lastRowB, "B").MergeArea
            lastRowB = mergedRange.Row + mergedRange.Rows.Count - 1
        End If

        ' Use the maximum of both columns
        lastRow = WorksheetFunction.Max(lastRow, lastRowB)
    End If

    ' Start adding dates in the next row
    currentRow = lastRow + 1
    yearMergeStartRow = currentRow
    monthMergeStartRow = currentRow
    previousYear = ""
    previousMonth = ""

    ' Loop through each day of the year
    For currentDate = startDate To DateAdd("yyyy", 1, startDate) - 1
        Dim currentYear As String
        Dim currentMonth As String

        currentYear = Format(currentDate, "yy")
        currentMonth = Format(currentDate, "mmm")

        ' Write the date components
        ws.Cells(currentRow, 1) = currentYear     ' Year (YY)
        ws.Cells(currentRow, 2) = currentMonth    ' Month (MMM)
        ws.Cells(currentRow, 3) = Format(currentDate, "dd")     ' Day (DD)
        ws.Cells(currentRow, 4) = Format(currentDate, "ddd")    ' Day abbreviation (DDD)

        ' Check if year changed
        If currentYear <> previousYear And currentRow > yearMergeStartRow Then
            ' Merge year cells
            If currentRow - yearMergeStartRow > 0 Then
                ws.Range(ws.Cells(yearMergeStartRow, 1), ws.Cells(currentRow - 1, 1)).Merge
            End If
            yearMergeStartRow = currentRow
        End If

        ' Check if month changed
        If currentMonth <> previousMonth And currentRow > monthMergeStartRow Then
            ' Merge month cells
            If currentRow - monthMergeStartRow > 0 Then
                ws.Range(ws.Cells(monthMergeStartRow, 2), ws.Cells(currentRow - 1, 2)).Merge
            End If
            monthMergeStartRow = currentRow
        End If

        previousYear = currentYear
        previousMonth = currentMonth

        ' Move to next row
        currentRow = currentRow + 1
    Next currentDate

    ' Merge the last groups
    If currentRow - yearMergeStartRow > 0 Then
        ws.Range(ws.Cells(yearMergeStartRow, 1), ws.Cells(currentRow - 1, 1)).Merge
    End If
    If currentRow - monthMergeStartRow > 0 Then
        ws.Range(ws.Cells(monthMergeStartRow, 2), ws.Cells(currentRow - 1, 2)).Merge
    End If

    ' Format and align with all borders
    With ws.Range(ws.Cells(lastRow + 1, 1), ws.Cells(currentRow - 1, 4))
        ' Clear any existing borders first
        .Borders.LineStyle = xlNone

        ' Set all borders
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin

            ' Ensure all border sides are set
            .Item(xlEdgeLeft).LineStyle = xlContinuous
            .Item(xlEdgeTop).LineStyle = xlContinuous
            .Item(xlEdgeBottom).LineStyle = xlContinuous
            .Item(xlEdgeRight).LineStyle = xlContinuous
            .Item(xlInsideVertical).LineStyle = xlContinuous
            .Item(xlInsideHorizontal).LineStyle = xlContinuous
        End With
    End With


    ' Set top-left alignment for merged cells
    With ws.Range(ws.Cells(lastRow + 1, 1), ws.Cells(currentRow - 1, 2))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With

    ' Center align the day and day abbreviation columns
    With ws.Range(ws.Cells(lastRow + 1, 3), ws.Cells(currentRow - 1, 4))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Auto-fit columns
    ws.Columns("A:D").AutoFit

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Dates populated successfully!", vbInformation
End Sub
