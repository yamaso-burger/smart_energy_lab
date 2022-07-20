Attribute VB_Name = "Module4"
Option Explicit

Sub 積算平均_共鳴信号真数_2022年_5月改良()
    Dim measurementPoints As Integer
    Dim amountOfData As Integer
    Dim measurementTimes As Integer
    
    Dim isExecute As Boolean
    Dim isAverage As Boolean
    
    Dim test As Double
    
    
    If Not Cells(2, 1) = "" Then
        isExecute = True
        
        
        measurementPoints = Cells(3, 2).End(xlDown).Row
        '積算平均結果を取り除く
        If Not Columns(1).Find("積算平均結果") Is Nothing Then
            Dim ave_ini As Integer
            
            ave_ini = Cells(Rows.Count, 2).End(xlUp).End(xlUp).Row - 1
            Range(Cells(ave_ini, 1), Cells(ave_ini + measurementPoints + 1, 3)).ClearContents
            
            
        End If
            
        measurementPoints = Cells(3, 2).End(xlDown).Row
        amountOfData = Cells(Rows.Count, 1).End(xlUp).Row
        
        measurementTimes = amountOfData \ measurementPoints
        
        amountOfData = measurementTimes * measurementPoints
        
        
        If Columns(1).Find("積算平均結果") Is Nothing Then
        
            Dim i
            For i = 3 To measurementPoints
                Cells(i, 5) = Cells(i, 1)
                
                Cells(i, 6) = _
                WorksheetFunction.SumIf(Range(Cells(3, 1), Cells(amountOfData, 1)), Cells(i, 1), Range(Cells(3, 2), Cells(amountOfData, 2))) _
                / measurementTimes
                
                Cells(i, 7) = _
                WorksheetFunction.SumIf(Range(Cells(3, 1), Cells(amountOfData, 1)), Cells(i, 1), Range(Cells(3, 3), Cells(amountOfData, 3))) _
                / measurementTimes
                
                
                'Absorbanceを真数に変換する
                Cells(i, 17) = Cells(i, 5)
                Cells(i, 18) = Application.WorksheetFunction.Power(10, Cells(i, 6) / 20)
                Cells(i, 19) = Cells(i, 7)
                
            Next i
            isAverage = False
            
        Else
            MsgBox "copying data"
            amountOfData = amountOfData - measurementPoints
            Dim j
            Dim averageDataIndex As Integer
            averageDataIndex = (Cells(Rows.Count, 2).End(xlUp).Row + 1) - (measurementPoints - 3)
            
            For j = 3 To measurementPoints
            
                Cells(j, 5) = Cells(averageDataIndex, 1)
                Cells(j, 6) = Cells(averageDataIndex, 2)
                Cells(j, 7) = Cells(averageDataIndex, 3)
                averageDataIndex = averageDataIndex + 1
                
                'Absorbanceを真数に変換する
                Cells(j, 17) = Cells(j, 5)
                Cells(j, 18) = Application.WorksheetFunction.Power(10, Cells(j, 6) / 20)
                Cells(j, 19) = Cells(j, 7)
                
            Next j
            
            isAverage = True
        End If
        
        Dim minValue As Double
        minValue = WorksheetFunction.Min(Range(Cells(3, 18), Cells(measurementPoints, 18)))
        Dim fieldValueAtMin As Double
        fieldValueAtMin = WorksheetFunction.Index(Range(Cells(3, 17), Cells(measurementPoints, 17)), WorksheetFunction.Match(minValue, Range(Cells(3, 18), Cells(measurementPoints, 18)), 0))
       
        For i = 3 To measurementPoints
            Cells(i, 20) = (Cells(i, 17) - fieldValueAtMin) * 10
            Cells(i, 28) = Cells(i, 20)
            Cells(i, 32) = Cells(i, 20)
        Next i
        
        '//中央値を求めて起電力のオフセットを０にする
        Dim average_0 As Double
        
        Dim bottom_row_index
        bottom_row_index = Cells(Rows.Count, 17).End(xlUp).Row
        
        average_0 = WorksheetFunction.Average(Range(Cells(4, 19), Cells(14, 19)), Range(Cells(bottom_row_index - 1, 19), Cells(bottom_row_index - 11, 19)))
        '//磁界のオフセット０
        Dim absorbanceGround As Double
        absorbanceGround = WorksheetFunction.Average(Range(Cells(4, 18), Cells(14, 18)), Range(Cells(bottom_row_index - 1, 18), Cells(bottom_row_index - 11, 18)))
        
        Dim k
        
        For k = 3 To measurementPoints
                Cells(k, 21) = Cells(k, 19) - average_0
                Cells(k, 29) = Cells(k, 21)
                Cells(k, 33) = Cells(k, 18) - absorbanceGround
            Next k
        
    Else
        MsgBox "0 degree: error, paste data correctly!"
        isExecute = False
    End If
    
    Dim measurementPointsN As Integer
    Dim amountOfDataN As Integer
    Dim measurementTimesN As Integer
    
    Dim isAverageN As Boolean
    Dim isExecuteN As Boolean
    
    Dim N_fix As Integer
    N_fix = 8
    
    If Not Cells(2, 9) = "" Then
        isExecuteN = True
        measurementPointsN = Cells(3, 10).End(xlDown).Row
        
        '積算平均結果を取り除く
        If Not Columns(1 + N_fix).Find("積算平均結果") Is Nothing Then
            Dim ave_iniN As Integer
            
            ave_iniN = Cells(Rows.Count, 2 + N_fix).End(xlUp).End(xlUp).Row - 1
            Range(Cells(ave_iniN, 1 + N_fix), Cells(ave_iniN + measurementPoints + 1, 3 + N_fix)).ClearContents
            
            
        End If
        
        measurementPointsN = Cells(3, 10).End(xlDown).Row
        amountOfDataN = Cells(Rows.Count, 9).End(xlUp).Row
        
        measurementTimesN = amountOfDataN \ measurementPointsN
        
        amountOfDataN = measurementTimesN * measurementPointsN
        
        
        If Columns(9).Find("積算平均結果") Is Nothing Then
            For i = 3 To measurementPointsN
                Cells(i, 13) = Cells(i, 9)
                
                Cells(i, 14) = _
                WorksheetFunction.SumIf(Range(Cells(3, 9), Cells(amountOfDataN, 9)), Cells(i, 9), Range(Cells(3, 10), Cells(amountOfDataN, 10))) _
                / measurementTimesN
                
                Cells(i, 15) = _
                WorksheetFunction.SumIf(Range(Cells(3, 9), Cells(amountOfDataN, 9)), Cells(i, 9), Range(Cells(3, 11), Cells(amountOfDataN, 11))) _
                / measurementTimesN
                
                'Absorbanceを真数にする
                Cells(i, 23) = Cells(i, 13)
                Cells(i, 24) = Application.WorksheetFunction.Power(10, Cells(i, 14) / 20)
                Cells(i, 25) = Cells(i, 15)
            
            Next i
            isAverageN = False
            
        Else
            MsgBox "copying data"
            amountOfDataN = amountOfDataN - measurementPointsN
            Dim averageDataIndexN As Integer
            averageDataIndexN = (Cells(Rows.Count, 9).End(xlUp).Row + 1) - (measurementPointsN - 3)
            
            For j = 3 To measurementPointsN
            
                Cells(j, 13) = Cells(averageDataIndexN, 9)
                Cells(j, 14) = Cells(averageDataIndexN, 10)
                Cells(j, 15) = Cells(averageDataIndexN, 11)
                averageDataIndexN = averageDataIndexN + 1
                Cells(j, 23) = Cells(j, 13)
                Cells(j, 24) = Cells(j, 14)
                Cells(j, 25) = Cells(j, 15)
                
            Next j
            
            isAverageN = True
            
        End If
        
        Dim minValueN As Double
        minValueN = WorksheetFunction.Min(Range(Cells(3, 24), Cells(measurementPointsN, 24)))
        Dim fieldValueAtMinN As Double
        fieldValueAtMinN = WorksheetFunction.Index(Range(Cells(3, 23), Cells(measurementPointsN, 23)), WorksheetFunction.Match(minValueN, Range(Cells(3, 24), Cells(measurementPointsN, 24)), 0))
        For i = 3 To measurementPointsN
            Cells(i, 26) = (Cells(i, 23) - fieldValueAtMinN) * 10
            Cells(i, 30) = Cells(i, 26)
            Cells(i, 34) = Cells(i, 26)
        Next i
        
        Dim average_180 As Double
        
        '//起電力オフセット平均
        
        Dim bottom_row_indexN
        bottom_row_indexN = Cells(Rows.Count, 23).End(xlUp).Row
        
        average_180 = WorksheetFunction.Average(Range(Cells(4, 25), Cells(14, 25)), Range(Cells(bottom_row_indexN - 1, 25), Cells(bottom_row_indexN - 11, 25)))
        
         '//磁界のオフセット０
        Dim absorbanceGroundN As Double
        absorbanceGroundN = WorksheetFunction.Average(Range(Cells(4, 24), Cells(14, 24)), Range(Cells(bottom_row_index - 1, 24), Cells(bottom_row_index - 11, 24)))
        
        Dim l
        
        For l = 3 To measurementPointsN
                Cells(l, 27) = Cells(l, 25) - average_180
                Cells(l, 31) = Cells(l, 27)
                Cells(l, 35) = Cells(l, 24) - absorbanceGroundN
            Next l
        
        
    Else
        MsgBox "180 degree: error, paste data correctly!"
        isExecuteN = False
    End If
    
    
    
    
    
    
    
    
    
    If isExecute And isExecuteN Then
        MsgBox "処理完了>>" & vbCrLf & _
        "0°データ数:" & amountOfData & " , 測定回数:" & measurementTimes & _
        " , 測定ポイント数：" & measurementPoints - 3 & vbCrLf & _
        ">> 共鳴磁界値：" & fieldValueAtMin & vbCrLf & _
        ">> P(ground0) = " & absorbanceGround & vbCrLf & _
        ">> 180°データ数:" & amountOfDataN & " , 測定回数:" & measurementTimesN & _
        " , 測定ポイント数：" & measurementPointsN - 3 & vbCrLf & _
        ">> 共鳴磁界値：" & fieldValueAtMinN & vbCrLf & _
        ">> P(ground180) = " & absorbanceGroundN
    ElseIf isExecute Then
        MsgBox "処理完了>>" & vbCrLf & _
        "0°データ数:" & amountOfData & " , 測定回数:" & measurementTimes & _
        " , 測定ポイント数：" & measurementPoints - 3 & vbCrLf & _
        ">> 共鳴磁界値：" & fieldValueAtMin & vbCrLf & _
        ">> P(ground0) = " & absorbanceGround
    ElseIf isExecuteN Then
        MsgBox "処理完了>>" & vbCrLf & _
        ">> 180°データ数:" & amountOfDataN & " , 測定回数:" & measurementTimesN & _
        " , 測定ポイント数：" & measurementPointsN - 3 & vbCrLf & _
        ">> 共鳴磁界値：" & fieldValueAtMinN & vbCrLf & _
        ">> P(ground180) = " & absorbanceGroundN
    Else
        MsgBox "no data"
    End If
    

End Sub


