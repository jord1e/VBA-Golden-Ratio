' https://gamedev.stackexchange.com/questions/46463/how-can-i-find-an-optimum-set-of-colors-for-10-players
Public Sub ColorUniqueCellsUsingGoldenRatio(range As range)
    Dim dictionary As New Scripting.dictionary
    
    Dim i%
    For Each cel In range
        With cel
            If Not dictionary.Exists(.Value) Then
                dictionary.Add .Value, GoldenRatioColorFromIndex(i)
                i = i + 1
            End If
            .Interior.color = dictionary.Item(.Value)
        End With
    Next cel
End Sub
Public Function GoldenRatioColorFromIndex(index%) As Long
    Dim Hue#, Saturation#, Value#
    
    Hue = FMod(CDbl(index) * 0.618033988749895, 1#)
    Saturation = 0.5
    Value = Sqr(1# - FMod(CDbl(index) * 0.618033988749895, 0.5))
    
    GoldenRatioColorFromIndex = HSVtoRGB(Hue * 360, Saturation, Value)
End Function
' Copied From: https://www.howtobuildsoftware.com/index.php/how-do/cr3z/vba-ms-access-double-modulus-mod-with-doubles
Public Function FMod(a As Double, b As Double) As Double
    FMod = a - Fix(a / b) * b

    'http://en.wikipedia.org/wiki/Machine_epsilon
    'Unfortunately, this function can only be accurate when `a / b` is outside [-2.22E-16,+2.22E-16]
    'Without this correction, FMod(.66, .06) = 5.55111512312578E-17 when it should be 0
    If FMod >= -2 ^ -52 And FMod <= 2 ^ -52 Then '+/- 2.22E-16
        FMod = 0
    End If
End Function
' Adapted From: https://stackoverflow.com/questions/3018313/algorithm-to-convert-rgb-to-hsv-and-hsv-to-rgb-in-range-0-255-for-both
Public Function HSVtoRGB(Hue#, Saturation#, Value#) As Long
    Dim hh#, p#, q#, t#, ff#
    Dim i&
    
    If Saturation <= 0# Then
        Value = Value * 255
        HSVtoRGB = rgb(Value, Value, Value)
        Exit Function
    End If
    
    hh = IIf(Hue >= 360#, 0#, Hue) / 60#
    i = CLng(hh)
    ff = hh - i
    p = Value * (1# - Saturation)
    q = Value * (1# - (Saturation * ff))
    t = Value * (1# - (Saturation * (1# - ff)))
    
    Value = Value * 255
    p = p * 255
    t = t * 255
    q = q * 255
    Select Case i
        Case 0
            HSVtoRGB = rgb(Value, t, p)
        Case 1
            HSVtoRGB = rgb(q, Value, p)
        Case 2
            HSVtoRGB = rgb(p, Value, t)
        Case 3
            HSVtoRGB = rgb(p, q, Value)
        Case 4
            HSVtoRGB = rgb(t, p, Value)
        Case Else
            HSVtoRGB = rgb(Value, p, q)
    End Select
End Function

