Sub ProcessTest()

    Const ConfigSourceSheetName As String = ""

    Dim wsSource As Worksheet, wsData As Worksheet, wsSummary As Worksheet
    Dim wb As Workbook
    Dim r As Long, c As Long, i As Long, b As Long, k As Long, t As Long
    Dim headerRow As Long, dataStartRow As Long, lastRow As Long, maxCol As Long
    Dim blockCount As Long
    Dim blockStartCols() As Long, deviceIDs() As String, testNames() As String
    Dim rawVal As String, tsText As String
    Dim posTest As Long, wVal As Variant, elapsedSec As Double
    Dim outRow As Long, sumHeaderRow As Long, sumRow As Long
    Dim nRows As Long, rowIdx As Long
    Dim bMaxWeight() As Double, bStartRows() As Long
    Dim timeArr() As Double, weightArr() As Double
    Dim distinctTests() As String, distinctCount As Long
    Dim found As Boolean
    Dim chtObj As ChartObject
    Dim chartTop As Double, chartLeft As Double, chartWidth As Double, chartHeight As Double
    Dim gap As Double, baseTop As Double, baseLeft As Double
    Dim colIdx As Long, rowIdxGrid As Long
    Dim dataName As String, summaryName As String
    Dim srcData As Variant, masterData() As Variant
    Dim totalPossibleRows As Long, splitArr() As String
    Dim btn As Object
    
    Dim tsVal As Variant
    Dim firstFmt As String, isTimeFmt As Boolean
    Dim firstTime As Double, isFirst As Boolean, rawSec As Double

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set wb = ThisWorkbook
    If ConfigSourceSheetName = "" Then
        Set wsSource = ActiveSheet
    Else
        Set wsSource = wb.Worksheets(ConfigSourceSheetName)
    End If

    headerRow = 0
    For r = 1 To 30
        If Trim(wsSource.Cells(r, 1).Value) = "Timestamp" Then
            headerRow = r
            Exit For
        End If
    Next r
    If headerRow = 0 Then
        MsgBox "Could not find a 'Timestamp' header row in the first 30 rows of """ & wsSource.Name & """.", vbExclamation
        GoTo CleanExit
    End If
    dataStartRow = headerRow + 1

    maxCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    blockCount = 0
    ReDim blockStartCols(1 To maxCol)
    ReDim deviceIDs(1 To maxCol)
    ReDim testNames(1 To maxCol)

    c = 1
    Do While c <= maxCol
        If Trim(wsSource.Cells(1, c).Value) = "Device & Test:" Then
            rawVal = Trim(wsSource.Cells(1, c + 1).Value)
            posTest = InStr(rawVal, "_Test_")
            If posTest > 0 Then
                blockCount = blockCount + 1
                blockStartCols(blockCount) = c
                deviceIDs(blockCount) = Trim(Left(rawVal, posTest - 1))
                testNames(blockCount) = Trim(Mid(rawVal, posTest + Len("_Test_")))
            End If
            c = c + 3
        Else
            c = c + 1
        End If
    Loop

    If blockCount = 0 Then
        MsgBox "No 'Device & Test:' blocks found in row 1 of """ & wsSource.Name & """.", vbExclamation
        GoTo CleanExit
    End If

    lastRow = wsSource.Cells(wsSource.Rows.Count, blockStartCols(1)).End(xlUp).Row
    summaryName = Left(wsSource.Name & " - Summary", 31)
    dataName = Left(wsSource.Name & " - Data", 31)

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(summaryName).Delete
    wb.Worksheets(dataName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsSummary = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsSummary.Name = summaryName
    Set wsData = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsData.Name = dataName

    wsData.Range("A1:E1").Value = Array("Legend Name", "Chart Name", "Timestamp", "Elapsed Seconds", "Weight (lbf)")
    wsData.Range("A1:E1").Font.Bold = True
    wsData.Range("A1:E1").Interior.Color = RGB(220, 230, 241)

    totalPossibleRows = blockCount * (lastRow - dataStartRow + 1)
    ReDim masterData(1 To totalPossibleRows, 1 To 5)
    outRow = 2

    ReDim bMaxWeight(1 To blockCount)
    ReDim bStartRows(1 To blockCount)

    For b = 1 To blockCount
        srcData = wsSource.Range(wsSource.Cells(dataStartRow, blockStartCols(b)), wsSource.Cells(lastRow, blockStartCols(b) + 1)).Value2
        nRows = UBound(srcData, 1)
        ReDim timeArr(1 To nRows)
        ReDim weightArr(1 To nRows)
        rowIdx = 0
        bStartRows(b) = outRow
        
        isFirst = True
        firstFmt = wsSource.Cells(dataStartRow, blockStartCols(b)).NumberFormat
        isTimeFmt = (InStr(firstFmt, ":") > 0)

        For i = 1 To nRows
            tsVal = srcData(i, 1)
            If IsEmpty(tsVal) Or IsError(tsVal) Then Exit For
            
            tsText = Trim(CStr(tsVal))
            If tsText = "" Then Exit For
            wVal = srcData(i, 2)

            If isTimeFmt And VarType(tsVal) = vbDouble Then
                rawSec = tsVal * 86400
            Else
                tsText = Replace(Replace(tsText, " AM", "", 1, -1, vbTextCompare), " PM", "", 1, -1, vbTextCompare)
                If InStr(tsText, ":") > 0 Then
                    splitArr = Split(tsText, ":")
                    If UBound(splitArr) = 2 Then
                        rawSec = CDbl(splitArr(0)) * 3600 + CDbl(splitArr(1)) * 60 + CDbl(splitArr(2))
                    ElseIf UBound(splitArr) = 1 Then
                        rawSec = CDbl(splitArr(0)) * 60 + CDbl(splitArr(1))
                    Else
                        rawSec = CDbl(tsText)
                    End If
                Else
                    rawSec = CDbl(tsText)
                End If
            End If

            If isFirst Then
                firstTime = rawSec
                isFirst = False
            End If
            
            elapsedSec = rawSec - firstTime

            rowIdx = rowIdx + 1
            timeArr(rowIdx) = elapsedSec
            weightArr(rowIdx) = CDbl(wVal) * 0.00220462

            If rowIdx = 1 Or weightArr(rowIdx) > bMaxWeight(b) Then
                bMaxWeight(b) = weightArr(rowIdx)
            End If

            masterData(outRow - 1, 1) = deviceIDs(b)
            masterData(outRow - 1, 2) = testNames(b)
            masterData(outRow - 1, 3) = tsText
            masterData(outRow - 1, 4) = elapsedSec
            masterData(outRow - 1, 5) = weightArr(rowIdx)
            outRow = outRow + 1
        Next i
    Next b

    If outRow > 2 Then
        wsData.Range("A2").Resize(outRow - 2, 5).Value = masterData
    End If

    wsData.Range("D2:D" & outRow - 1).NumberFormat = "0.000"
    wsData.Columns("A:E").AutoFit
    wsData.Activate
    wsData.Range("A2").Select
    ActiveWindow.FreezePanes = True

    wsSummary.Activate
    wsSummary.Range("A1").Value = "Test Summary - " & wsSource.Name
    wsSummary.Range("A1").Font.Bold = True
    wsSummary.Range("A1").Font.Size = 14
    
    wsSummary.Range("A2").Value = "Rolling Avg Window (s):"
    wsSummary.Range("A2").Font.Italic = True
    wsSummary.Range("B2").Value = 0
    wsSummary.Range("B2").Font.Bold = True
    wsSummary.Range("B2").Interior.Color = RGB(255, 242, 204)
    wsSummary.Range("B2").Borders.LineStyle = xlContinuous
    wsSummary.Range("B2").HorizontalAlignment = xlCenter

    Set btn = wsSummary.Buttons.Add(wsSummary.Range("D2").Left, wsSummary.Range("D2").Top - 2, 100, 22)
    btn.OnAction = "UpdateAverages"
    btn.Characters.Text = "Update Charts"
    btn.Font.Bold = True

    sumHeaderRow = 4
    wsSummary.Cells(sumHeaderRow, 1).Resize(1, 4).Value = Array("Legend Name", "Chart Name", "Raw Peak (lbf)", "Filtered Peak (lbf)")
    wsSummary.Cells(sumHeaderRow, 1).Resize(1, 4).Font.Bold = True
    wsSummary.Cells(sumHeaderRow, 1).Resize(1, 4).Interior.Color = RGB(220, 230, 241)

    sumRow = sumHeaderRow + 1
    For b = 1 To blockCount
        wsSummary.Cells(sumRow, 1).Value = deviceIDs(b)
        wsSummary.Cells(sumRow, 2).Value = testNames(b)
        wsSummary.Cells(sumRow, 3).Value = bMaxWeight(b)
        wsSummary.Cells(sumRow, 4).Value = bMaxWeight(b)
        sumRow = sumRow + 1
    Next b

    wsSummary.Range("C" & (sumHeaderRow + 1) & ":D" & (sumRow - 1)).NumberFormat = "0.00"
    wsSummary.Range("A" & sumHeaderRow & ":D" & (sumRow - 1)).Borders.LineStyle = xlContinuous
    wsSummary.Columns("A:D").AutoFit

    distinctCount = 0
    ReDim distinctTests(1 To blockCount)
    For b = 1 To blockCount
        found = False
        For k = 1 To distinctCount
            If distinctTests(k) = testNames(b) Then found = True: Exit For
        Next k
        If Not found Then
            distinctCount = distinctCount + 1
            distinctTests(distinctCount) = testNames(b)
        End If
    Next b

    chartWidth = 380
    chartHeight = 250
    gap = 15
    baseLeft = wsSummary.Cells(1, 6).Left
    baseTop = wsSummary.Cells(1, 1).Top

    For t = 1 To distinctCount
        colIdx = (t - 1) Mod 2
        rowIdxGrid = (t - 1) \ 2
        chartLeft = baseLeft + colIdx * (chartWidth + gap)
        chartTop = baseTop + rowIdxGrid * (chartHeight + gap)

        Set chtObj = wsSummary.ChartObjects.Add(Left:=chartLeft, Top:=chartTop, Width:=chartWidth, Height:=chartHeight)
        chtObj.chart.ChartType = xlXYScatter
        chtObj.chart.HasTitle = True
        chtObj.chart.ChartTitle.Text = distinctTests(t)
        chtObj.chart.Legend.Position = xlLegendPositionBottom
        chtObj.chart.Axes(xlCategory).HasTitle = True
        chtObj.chart.Axes(xlCategory).AxisTitle.Text = "Elapsed Seconds"
        chtObj.chart.Axes(xlValue).HasTitle = True
        chtObj.chart.Axes(xlValue).AxisTitle.Text = "Weight (lbf)"
    Next t

    wsSummary.Range("A1").Select
    Call UpdateAverages

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


Sub UpdateAverages()
    On Error GoTo PlottingError

    Dim wsSum As Worksheet, wsDat As Worksheet
    Set wsSum = ActiveSheet
    
    If InStr(wsSum.Name, " - Summary") = 0 Then Exit Sub
    
    Set wsDat = ThisWorkbook.Sheets(Replace(wsSum.Name, " - Summary", " - Data"))
    If wsDat Is Nothing Then Exit Sub
    
    Dim avgWindow As Double
    avgWindow = Val(wsSum.Range("B2").Value)
    
    Application.ScreenUpdating = False
    
    wsDat.Range("G:XFD").ClearContents
    
    Dim lastR As Long
    lastR = wsDat.Cells(wsDat.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then GoTo SafeExit
    
    Dim rawData As Variant
    rawData = wsDat.Range("A1:E" & lastR).Value2
    
    Dim chrtObj As ChartObject, ser As Series
    Dim tName As String, testNameStr As String, dName As String
    Dim outCol As Long
    outCol = 7
    
    For Each chrtObj In wsSum.ChartObjects
        tName = chrtObj.Chart.ChartTitle.Text
        If InStr(tName, " (") > 0 Then
            testNameStr = Trim(Left(tName, InStrRev(tName, " (") - 1))
        Else
            testNameStr = Trim(tName)
        End If
        
        chrtObj.Chart.ChartTitle.Text = testNameStr & IIf(avgWindow > 0, " (" & avgWindow & "s Avg)", " (Raw)")
        
        Do While chrtObj.Chart.SeriesCollection.Count > 0
            chrtObj.Chart.SeriesCollection(1).Delete
        Loop
        
        Dim uDevs As Collection
        Set uDevs = New Collection
        Dim i As Long
        On Error Resume Next
        For i = 2 To lastR
            If Trim(CStr(rawData(i, 2))) = testNameStr Then
                uDevs.Add CStr(rawData(i, 1)), CStr(rawData(i, 1))
            End If
        Next i
        On Error GoTo PlottingError
        
        Dim devKey As Variant
        For Each devKey In uDevs
            dName = CStr(devKey)
            
            Dim matchCount As Long
            matchCount = 0
            For i = 2 To lastR
                If Trim(CStr(rawData(i, 2))) = testNameStr And CStr(rawData(i, 1)) = dName Then
                    matchCount = matchCount + 1
                End If
            Next i
            
            If matchCount > 0 Then
                Dim tArr() As Double, wArr() As Double
                ReDim tArr(1 To matchCount)
                ReDim wArr(1 To matchCount)
                Dim idx As Long
                idx = 0
                
                For i = 2 To lastR
                    If Trim(CStr(rawData(i, 2))) = testNameStr And CStr(rawData(i, 1)) = dName Then
                        idx = idx + 1
                        tArr(idx) = CDbl(rawData(i, 4))
                        wArr(idx) = CDbl(rawData(i, 5))
                    End If
                Next i
                
                Set ser = chrtObj.Chart.SeriesCollection.NewSeries
                ser.Name = dName
                
                Dim outData() As Variant
                ReDim outData(1 To matchCount, 1 To 2)
                
                Dim filteredMax As Double
                filteredMax = -999999
                
                If avgWindow > 0 Then
                    Dim headIdx As Long, tailIdx As Long
                    Dim rollingSum As Double
                    Dim rollingCount As Long
                    Dim currentAvg As Double
                    
                    tailIdx = 1
                    rollingSum = 0
                    
                    For headIdx = 1 To matchCount
                        rollingSum = rollingSum + wArr(headIdx)
                        
                        Do While (tArr(headIdx) - tArr(tailIdx) > avgWindow) And (tailIdx <= headIdx)
                            rollingSum = rollingSum - wArr(tailIdx)
                            tailIdx = tailIdx + 1
                        Loop
                        
                        rollingCount = headIdx - tailIdx + 1
                        
                        outData(headIdx, 1) = tArr(headIdx)
                        If rollingCount > 0 Then
                            currentAvg = rollingSum / rollingCount
                        Else
                            currentAvg = wArr(headIdx)
                        End If
                        
                        outData(headIdx, 2) = currentAvg
                        If currentAvg > filteredMax Then filteredMax = currentAvg
                    Next headIdx
                    
                Else
                    For i = 1 To matchCount
                        outData(i, 1) = tArr(i)
                        outData(i, 2) = wArr(i)
                        If wArr(i) > filteredMax Then filteredMax = wArr(i)
                    Next i
                End If
                
                wsDat.Cells(1, outCol).Value = dName & " Time"
                wsDat.Cells(1, outCol + 1).Value = dName & " lbf"
                wsDat.Cells(2, outCol).Resize(matchCount, 2).Value = outData
                
                ser.XValues = wsDat.Range(wsDat.Cells(2, outCol), wsDat.Cells(matchCount + 1, outCol))
                ser.Values = wsDat.Range(wsDat.Cells(2, outCol + 1), wsDat.Cells(matchCount + 1, outCol + 1))
                outCol = outCol + 2
                
                ser.ChartType = xlXYScatter
                ser.MarkerStyle = xlMarkerStyleCircle
                ser.MarkerSize = 4
                
                Dim sumR As Long, maxSumR As Long
                maxSumR = wsSum.Cells(wsSum.Rows.Count, 1).End(xlUp).Row
                For sumR = 5 To maxSumR
                    If wsSum.Cells(sumR, 1).Value = dName And wsSum.Cells(sumR, 2).Value = testNameStr Then
                        wsSum.Cells(sumR, 4).Value = filteredMax
                        Exit For
                    End If
                Next sumR
                
            End If
        Next devKey
    Next chrtObj
    
SafeExit:
    Application.ScreenUpdating = True
    Exit Sub
    
PlottingError:
    Application.ScreenUpdating = True
    MsgBox "An error occurred while building the charts." & vbCrLf & "Error Code: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbCritical, "Chart Builder Failed"
End Sub
