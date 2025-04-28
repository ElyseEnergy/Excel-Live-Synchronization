Option Explicit



Private Type TimingInfo

    StartTime As Double

    EndTime As Integer

    description As Integer

End Type



Private timings() As TimingInfo

Private timingCount As Long



Private Const MIN_VALUE As Double = 100

Private Const MAX_VALUE As Double = 220

Private Const EPSILON As Double = 0.01

Private Const INITIAL_STEP As Double = 120

Private Const FINE_STEP As Double = 5

Private Const START_ROW As Long = 34

Private Const END_ROW As Long = 49

Private Const TARGET_ROW As Long = 89

Private Const START_COL As Long = 4  ' Column D

Private Const END_COL As Long = 15   ' Column O



Private Sub StartTimer(description As String)

    ReDim Preserve timings(0 To timingCount)

    With timings(timingCount)

        .StartTime = Timer

        .description = description

    End With
	Debug.print "coucou"

End Sub



Private Sub StopTimer()

    timings(timingCount).EndTime = Timer

    timingCount = timingCount + 1

End Sub



Private Sub DisplayTimings()

    Dim i As Long

    Dim msg As String

    msg = "Performance Analysis:" & vbNewLine & vbNewLine

    

    For i = 0 To timingCount - 1

        msg = msg & timings(i).description & ": " & _

              Format(timings(i).EndTime - timings(i).StartTime, "0.00") & " seconds" & vbNewLine

    Next i

    

    MsgBox msg, vbInformation

End Sub



Sub AdjustValuesAcrossColumns()

    On Error GoTo ErrorHandler

    ReDim timings(0)

    timingCount = 0

    StartTimer "Total Execution"

    

    Application.ScreenUpdating = False

    Application.Calculation = xlCalculationManual

    

    Dim targetValue As Variant

    targetValue = GetTargetValue()

    If targetValue = False Then Exit Sub

    

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Power mix price forecast")

    

    Dim col As Long, row As Long

    Dim currentValue As Double, tempValue As Double

    

    For col = START_COL To END_COL

        StartTimer "Processing Column " & Split(ws.Cells(1, col).Address, "$")(1)

        ' Initialize column to minimum values

        ws.Range(ws.Cells(START_ROW, col), ws.Cells(END_ROW, col)).Value = MIN_VALUE

        

        If Not AdjustColumnValues(ws, col, CDbl(targetValue)) Then

            MsgBox "Could not reach target value for column " & Split(ws.Cells(1, col).Address, "$")(1), vbExclamation

        End If

        StopTimer  ' Stop column timer

    Next col

    

CleanExit:

    StopTimer  ' Stop total execution timer

    Application.ScreenUpdating = True

    Application.Calculation = xlCalculationAutomatic

    DisplayTimings

    Exit Sub

    

ErrorHandler:

    MsgBox "Error " & Err.Number & ": " & Err.description, vbCritical

    Resume CleanExit

End Sub



Private Function GetTargetValue() As Variant

    Dim targetValue As Variant

    targetValue = Application.InputBox("Enter the Production Target value for each cell in row 89:", Type:=1)

    

    If targetValue = False Then

        GetTargetValue = False

        Exit Function

    End If

    

    If targetValue < 0 Or targetValue > MAX_VALUE * 2 Then

        MsgBox "Please enter a valid target value between 0 and " & MAX_VALUE * 2, vbExclamation

        GetTargetValue = False

        Exit Function

    End If

    

    GetTargetValue = targetValue

End Function



Private Function AdjustColumnValues(ws As Worksheet, col As Long, targetValue As Double) As Boolean

    Dim row As Long, tempValue As Double, currentValue As Double

    

    For row = START_ROW To END_ROW

        If TryFindValue(ws, row, col, targetValue) Then

            AdjustColumnValues = True

            Exit Function

        End If

    Next row

    

    AdjustColumnValues = False

End Function



Private Function TryFindValue(ws As Worksheet, row As Long, col As Long, targetValue As Double) As Boolean

    Dim low As Double, high As Double, mid As Double

    Dim currentValue As Double

    Dim bestValue As Double

    Dim bestDiff As Double

    

    low = MIN_VALUE

    high = MAX_VALUE

    bestDiff = MAX_VALUE

    bestValue = low

    

    ' Binary search

    Do While high - low > FINE_STEP

        mid = (low + high) / 2

        ws.Cells(row, col).Value = mid

        ws.Calculate

        currentValue = ws.Cells(TARGET_ROW, col).Value

        

        ' Update best value if closer to target

        If Abs(currentValue - targetValue) < bestDiff Then

            bestDiff = Abs(currentValue - targetValue)

            bestValue = mid

            If bestDiff <= EPSILON Then

                TryFindValue = True

                Exit Function

            End If

        End If

        

        If currentValue < targetValue Then

            low = mid

        Else

            high = mid

        End If

    Loop

    

    ' Fine-tune around best value found

    Dim fineValue As Double

    For fineValue = bestValue - FINE_STEP To bestValue + FINE_STEP Step FINE_STEP / 4

        If fineValue >= MIN_VALUE And fineValue <= MAX_VALUE Then

            ws.Cells(row, col).Value = fineValue

            ws.Calculate

            currentValue = ws.Cells(TARGET_ROW, col).Value

            

            If Abs(currentValue - targetValue) <= EPSILON Then

                TryFindValue = True

                Exit Function

            End If

        End If

    Next fineValue

    

    ' Revert to best value found if no exact match

    ws.Cells(row, col).Value = bestValue

    TryFindValue = (bestDiff <= EPSILON)

End Function

