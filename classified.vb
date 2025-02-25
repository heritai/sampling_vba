' This function is to be used when dealing with stratified sampling techniques.
' it receives the ranges, do all math and it will give as Double, the final result (mean or variance).
' It will use a helper function called mult, created at the end.
Function mult(arr1() As Variant, arr2() As Variant) As Variant
    
    'Declare i variable as long for setting maximum array values
    Dim i As Long
    Dim res
    ReDim res(1 To UBound(arr1))
    
    'Setting iteration for doing calculus. It is the base for calculating mult array
    For i = 1 To UBound(arr1)
        res(i) = arr1(i) * arr2(i)
    Next i
    mult = res
End Function

' The main function that will calculate the end result that we need with a sample from population
Function cls(R1 As Range, R2 As Range, what As Byte) As Double

    'Definition of variants for calculating the parameters
    Dim Ar1 As Variant, Ar2 As Variant
    Ar1 = R1
    Ar2 = R2
    Dim ubnd As Long
    Dim s As Double
    Dim t As Double
    Dim num As Double
    
    'definition of the dimension
    Dim class()
    Dim avrgs()
    Dim vars()
    Dim weights()
    Dim counts()
    ubnd = UBound(Ar2, 1)
    
    ReDim class(1 To ubnd)
    ReDim arvgs(1 To ubnd)
    ReDim vars(1 To ubnd)
    ReDim weights(1 To ubnd)
    ReDim counts(1 To ubnd)
    
    'classes
    For j = 1 To UBound(Ar2, 1)
        class(j) = Ar2(j, 1)
    Next j
    
    'weights
    For t = 1 To UBound(Ar2, 1)
        weights(t) = Ar2(t, 3)
    Next t
    
    
    'Setting for what class the iterations will begin
    For k = 1 To UBound(class)
        s = 0
        num = 0
        For i = 1 To UBound(Ar1, 1)
            If Ar1(i, 1) = class(k) Then
                s = s + Ar1(i, 2)
                num = num + 1
            End If
        Next i
        arvgs(k) = s / num
        counts(k) = num
    Next k
    
    
    'Setting calculations by iterations
    For k = 1 To UBound(class)
        s = 0
        num = 0
        For i = 1 To UBound(Ar1, 1)
            If Ar1(i, 1) = class(k) Then
                s = s + (Ar1(i, 2) - arvgs(k)) ^ 2
                num = num + 1
            End If
        Next i
        vars(k) = s / (num - 1)
    Next k
    
    'MsgBox counts(4)
    t = 0
    
    'calculation using weights, in order to check the var
    For k = 1 To UBound(class)
        t = t + (weights(k) ^ 2) * (1 / counts(k) - 1 / Ar2(k, 2)) * vars(k)
    Next k
    
    
    'Here we have two cases where, if we check if it is what == 1 (then is calculating the average.)
    'but if it is not, will do the calculation for var.
    If what = 1 Then
        cls = WorksheetFunction.Sum(mult(weights, arvgs))
    Else
        cls = t
    End If
End Function