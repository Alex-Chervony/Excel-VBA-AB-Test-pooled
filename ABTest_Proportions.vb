Function ABTest_Proportions(n1, y1, n2, y2 As Long, Optional confidence As Double = 95) As Double
    ' This is a two tailed non pooled proportion test.
    
    Dim p1, p2, CIUpper, CILower As Double
    Dim significance As String
    
    'Proportions
    p1 = y1 / n1
    p2 = y2 / n2
    
    'Phat
    Phat = (y1 + y2) / (n1 + n2)
    
    If confidence = 95 Then
        'Two tailed test means we need -+97.5%
        pz = 1.959964
    ElseIf confidence = 90 Then
        'Two tailed test means we need -+95.0%
        pz = 1.644854
    Else
        'Proportion difference can be between -1 and 1
        ABTest_Proportions = -2
        Exit Function
    End If
    
    'Check if the test is applicable:
    'n1*p1 > 5 and n1*(1 - p1) > 5 and n2 * p2 > 5 and n2(1 - p2) > 5
    If (n1 * p1) > 5 And n1 * (1 - p1) > 5 And n2 * p2 > 5 And n2 * (1 - p2) > 5 And Frq1 = Frq2 Then
        'Z
        Z = (p1 - p2) / (Sqr((Phat * (1 - Phat)) * (1 / n1 + 1 / n2)))
        '95% Confidence Interval
        'CIUpper = 1.959964 * (Sqr((Phat * (1 - Phat)) * (1 / n1 + 1 / n2))) + p1 - p2
        'CILower = -1.959964 * (Sqr((Phat * (1 - Phat)) * (1 / n1 + 1 / n2))) + p1 - p2
        
        'Return proportion difference if the test has significant results
        If Abs(Z) > pz Then
            ABTest_Proportions = p1 - p2
            'Debug
            'ABTest_Proportions = Z
            Exit Function
        End If
        
    Else
        'Proportion difference can be between -1 and 1
        ABTest_Proportions = -2
        Exit Function

    End If
    
    ABTest_Proportions = -2
    
End Function
