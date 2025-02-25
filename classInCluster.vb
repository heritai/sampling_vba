' This function calculates the mean or variance for stratified sampling inside a class.
' This sampling method may use a class with the name 'classInClss'
' It receives the ranges required and calculates the mean/variance for each parameter

Function clssInClts(R1 As Range, R2 As Range, R3 As Range, N As Double, M As Double, what As Byte) As Double

    'Declare variables of Variant type to store a large number of data
    Dim Ar1 As Variant, Ar2 As Variant, Ar3 As Variant
    Ar1 = R1
    Ar2 = R2
    Ar3 = R3
    
    'Arrays declaration
    Dim avrgs()
    Dim vars()
    Dim counts()
    Dim LiNbarYbar()
    
    'Variable declaration
    Dim s As Double
    Dim v As Double
    Dim s2b As Double
    Dim t As Double
    Dim c As Long
    Dim ubnd As Long
    Dim clubnd As Long
    
    'Defines the maximum limits of parameters
    ubnd = UBound(Ar2, 1)
    clubnd = UBound(Ar3, 1)
    'Definition of the dimension
    ReDim avrgs(1 To ubnd)
    ReDim vars(1 To ubnd)
    ReDim counts(1 To ubnd)
    ReDim LiNbarYbar(1 To clubnd)
    
    
    'In this iteration, we loop to get data for each parameter
    For i = 1 To UBound(Ar2, 1)
        s = 0
        c = 0
        For j = 1 To UBound(Ar1, 1)
            If Ar1(j, 1) = Ar2(i, 1) And Ar1(j, 2) = Ar2(i, 2) Then
                s = s + Ar1(j, 3)
                c = c + 1
            End If
        Next j
        avrgs(i) = s / c
        counts(i) = c
    Next i
    
    
    'In this step, we iterate from the Ar2 parameters set in the previous step
    For i = 1 To UBound(Ar2, 1)
        v = 0
        For j = 1 To UBound(Ar1, 1)
            If Ar1(j, 1) = Ar2(i, 1) And Ar1(j, 2) = Ar2(i, 2) Then
                v = v + (Ar1(j, 3) - avrgs(i)) ^ 2
            End If
        Next j
        vars(i) = v / (counts(i) - 1)
    Next i
    
    
    'Calculation using classes iteration, where we loop from the class and calculate all metrics
    For k = 1 To UBound(Ar3, 1)
        s = 0
        For i = 1 To UBound(Ar2, 1)
            If Ar2(i, 1) = Ar3(k, 1) Then
                s = s + Ar2(i, 3) / Ar3(k, 2) * avrgs(i)
            End If
        Next i
        LiNbarYbar(k) = Ar3(k, 2) / (N / M) * s
    Next k
    
    
    'Calculation after previous steps
    t = 0
    For k = 1 To UBound(Ar3, 1)
        s = 0
        For i = 1 To UBound(Ar2, 1)
            If Ar2(i, 1) = Ar3(k, 1) Then
                s = s + (Ar2(i, 3) / Ar3(k, 2)) ^ 2 * (1 / counts(i) - 1 / Ar2(i, 3)) * vars(i)
            End If
        Next i
        t = t + (Ar3(k, 2) / (N / M)) ^ 2 * s
    Next k
    
    
    'With this command, we calculate the variance, calculating all Var parameters
    s2b = Application.WorksheetFunction.Var_S(LiNbarYbar)
    
    'Finalization
    If what = 1 Then
        clssInClts = Application.WorksheetFunction.Average(LiNbarYbar)
    Else
        clssInClts = (1 / UBound(Ar3, 1) - 1 / M) * s2b + 1 / (UBound(Ar3, 1) * M) * t
    End If
    
End Function