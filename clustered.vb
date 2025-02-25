' This function calculates the mean or variance for CLUSTERED sampling.
' It receives the ranges required and calculates the mean/variance for each parameter

Function clts(R1 As Range, R2 As Range, N As Long, m As Long, what As Byte) As Double

    'Declare variables of Variant type to store a large number of data
    Dim Ar1 As Variant, Ar2 As Variant
    Ar1 = R1
    Ar2 = R2
    
    'Arrays declaration
    Dim clusts()
    Dim avrgs()
    Dim Warvgs()
    Dim vars()
    Dim counts()
    
    'Variable declaration
    Dim s As Long
    Dim num As Long
    Dim ubnd As Long
    
    'Defines the maximum limits of parameters
    ubnd = UBound(Ar2, 1)
    'Definition of the dimension
    ReDim clusts(1 To ubnd)
    ReDim arvgs(1 To ubnd)
    ReDim Warvgs(1 To ubnd)
    ReDim vars(1 To ubnd)
    ReDim counts(1 To ubnd)
    
    'In this iteration, we loop to get data for each parameter
    For j = 1 To UBound(Ar2, 1)
        clusts(j) = Ar2(j, 1)
    Next j
    
    'In this step, we iterate from the clusters and calculate metrics
    For k = 1 To UBound(clusts)
        s = 0
        num = 0
        For i = 1 To UBound(Ar1, 1)
            If Ar1(i, 1) = clusts(k) Then
                s = s + Ar1(i, 2)
                num = num + 1
            End If
        Next i
        arvgs(k) = s / num
        counts(k) = num
    Next k
    
    'Calculations using averages iteration
    For k = 1 To UBound(clusts)
        s = 0
        num = 0
        For i = 1 To UBound(Ar1, 1)
            If Ar1(i, 1) = clusts(k) Then
                s = s + (Ar1(i, 2) - arvgs(k)) ^ 2
                num = num + 1
            End If
        Next i
        vars(k) = s / (num - 1)
    Next k
    
    'Calculations done for some metrics of our problem
    t = 0
    For k = 1 To UBound(clusts)
      t = t + Ar2(k, 2) / (N / m) * arvgs(k)
    Next k
    
    For k = 1 To UBound(clusts)
      Warvgs(k) = Ar2(k, 2) / (N / m) * arvgs(k)
    Next k
    
    
    'Calculate all var metrics with iterations
    tt = 0
    For k = 1 To UBound(clusts)
      tt = tt + ((Ar2(k, 2) / (N / m)) ^ 2) * (1 / counts(k) - 1 / Ar2(k, 2)) * vars(k)
    Next k
    
    
    'Finalization
    If what = 1 Then
        clts = t / UBound(clusts)
    Else
        clts = (1 / UBound(clusts) - 1 / m) * WorksheetFunction.Var_S(Warvgs) + (1 / (m * UBound(clusts))) * tt
    End If
End Function