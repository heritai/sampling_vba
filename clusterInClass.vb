' This function calculates the mean or variance for cluster inside class sampling.
' It receives the ranges required and calculates the mean/variance for each parameter

Function cltsInClss(R1 As Range, R2 As Range, R3 As Range, what As Byte) As Double

    'Definition of variants for calculating the parameters
    Dim Ar1 As Variant, Ar2 As Variant, Ar3 As Variant
    Ar1 = R1
    Ar2 = R2
    Ar3 = R3
    
    'Definition of the dimension
    Dim clusts()
    Dim class()
    Dim avrgs()
    Dim vars()
    Dim counts()
    Dim clustCount()
    Dim wMeans()
    
    'Declares variables of Double/Long type for calculating
    Dim s As Double
    Dim h As Double
    Dim vv As Double
    Dim s2b As Double
    Dim t As Double
    Dim d As Long
    Dim ss As Double
    Dim num As Long
    Dim ubnd As Long
    
    ubnd = UBound(Ar2, 1)
    
    'clusters
    ReDim clusts(1 To ubnd)
    For j = 1 To UBound(Ar2, 1)
        clusts(j) = Ar2(j, 2)
    Next j
    
    ReDim class(1 To 4)
    
    
    'classes
    Dim z As Integer
    For z = 1 To 4
      class(z) = Ar3(z, 1)
    Next z
    'class = Ar3(3, 1)
    
    
    ReDim clustCount(1 To 4)
    
    
    'cluster counts
    For j = 1 To 4
        d = 0
        For m = 1 To UBound(clusts)
           If Ar2(m, 1) = class(j) Then
            d = d + 1
           End If
        Next m
        clustCount(j) = d
    Next j
    
    
    'Setting for what class the iterations will begin
    For k = 1 To UBound(clusts)
        s = 0
        num = 0
        For i = 1 To UBound(Ar1, 1)
            If Ar1(i, 2) = clusts(k) Then
                s = s + Ar1(i, 3)
                num = num + 1
            End If
        Next i
        avrgs(k) = s / num
        counts(k) = num
    Next k
    
    
    'Calculations using averages iteration
    For k = 1 To UBound(clusts)
        s = 0
        num = 0
        For i = 1 To UBound(Ar1, 1)
            If Ar1(i, 2) = clusts(k) Then
                s = s + (Ar1(i, 3) - avrgs(k)) ^ 2
                num = num + 1
            End If
        Next i
        vars(k) = s / (num - 1)
    Next k
    
    v = 0
    For p = 1 To UBound(Ar3, 1)
      v = v + Ar3(p, 3)
    Next p
    
    ss = 0
    For k = 1 To 4
        t = 0
        For u = 1 To UBound(clusts)
            If Ar2(u, 1) = Ar3(k, 1) Then
            t = t + Ar2(u, 3) / (Ar3(k, 3) / Ar3(k, 2)) * avrgs(u)
            End If
        Next u
        ss = ss + (Ar3(k, 3) / v) / clustCount(k) * t
        
    Next k
    
     vv = 0
    For k = 1 To 4
        ReDim wMeans(1 To clustCount(k))
        ind = 0
        h = 0
        For u = 1 To UBound(clusts)
            If Ar2(u, 1) = Ar3(k, 1) Then
            ind = ind + 1
            wMeans(ind) = Ar2(u, 3) / (Ar3(k, 3) / Ar3(k, 2)) * avrgs(u)
            h = h + (Ar2(u, 3) / (Ar3(k, 3) / Ar3(k, 2))) ^ 2 * (1 / counts(u) - 1 / Ar2(u, 3)) * vars(u)
            End If
        Next u
        s2b = Application.WorksheetFunction.Var_S(wMeans)
        vv = vv + (Ar3(k, 3) / v) ^ 2 * ((1 / clustCount(k) - 1 / Ar3(k, 2)) * s2b + 1 / (clustCount(k) * Ar3(k, 2)) * h)
        
    Next k
    
    If what = 1 Then
        cltsInClss = ss
    Else
        cltsInClss = vv
    End If
End Function