Attribute VB_Name = "Module1"
Option Explicit
Function Vector(VectorRange As Range) As Variant

Dim List(1 To 3) As Double


If VectorRange.Columns.Count = 1 Then
    List(1) = VectorRange.Cells(1, 1).Value
    List(2) = VectorRange.Cells(2, 1).Value
    List(3) = VectorRange.Cells(3, 1).Value
Else
    List(1) = VectorRange.Cells(1, 1).Value
    List(2) = VectorRange.Cells(1, 2).Value
    List(3) = VectorRange.Cells(1, 3).Value
End If

Vector = List

End Function

Function VLen(VectorList As Variant) As Double

Dim X As Double
Dim Y As Double
Dim Z As Double

X = VectorList(1)
Y = VectorList(2)
Z = VectorList(3)

VLen = ((X ^ 2) + (Y ^ 2) + (Z ^ 2)) ^ (0.5)

End Function

Function DotProd(V1 As Variant, V2 As Variant) As Variant

Dim X1 As Double
Dim Y1 As Double
Dim Z1 As Double

Dim X2 As Double
Dim Y2 As Double
Dim Z2 As Double

X1 = V1(1)
Y1 = V1(2)
Z1 = V1(3)

X2 = V2(1)
Y2 = V2(2)
Z2 = V2(3)

DotProd = X1 * X2 + Y1 * Y2 + Z1 * Z2

End Function

Function XProd(V1 As Variant, V2 As Variant, Optional Index As Integer) As Variant

Dim X1 As Double
Dim Y1 As Double
Dim Z1 As Double

Dim X2 As Double
Dim Y2 As Double
Dim Z2 As Double

Dim Cx As Double
Dim Cy As Double
Dim Cz As Double

Dim XProdArr(1 To 3) As Double

X1 = V1(1)
Y1 = V1(2)
Z1 = V1(3)

X2 = V2(1)
Y2 = V2(2)
Z2 = V2(3)

Cx = Y1 * Z2 - Z1 * Y2
Cy = Z1 * X2 - X1 * Z2
Cz = X1 * Y2 - Y1 * X2

If Index > 0 Then
    Select Case Index
        Case 1
            XProd = Cx
        Case 2
            XProd = Cy
        Case 3
            XProd = Cz
    End Select
Else
    XProdArr(1) = Cx
    XProdArr(2) = Cy
    XProdArr(3) = Cz
    XProd = XProdArr
End If

End Function

Function VRotYZ(VectorList As Variant, gamma As Double, Optional Index As Integer) As Variant

Dim X As Double
Dim Y As Double
Dim Z As Double

Dim VRotArray(1 To 3) As Double

X = VectorList(1)
Y = VectorList(2)
Z = VectorList(3)

gamma = WorksheetFunction.Radians(gamma)

VRotArray(1) = X
VRotArray(2) = Cos(gamma) * Y - Sin(gamma) * Z
VRotArray(3) = Sin(gamma) * Y + Cos(gamma) * Z

If Index > 0 Then
    Select Case Index
        Case 1
            VRotYZ = VRotArray(1)
        Case 2
            VRotYZ = VRotArray(2)
        Case 3
            VRotYZ = VRotArray(3)
    End Select
Else
    VRotYZ = VRotArray
End If

End Function

Function VRotXY(VectorList As Variant, gamma As Double, Optional Index As Integer) As Variant

Dim X As Double
Dim Y As Double
Dim Z As Double

Dim VRotArray(1 To 3) As Double

X = VectorList(1)
Y = VectorList(2)
Z = VectorList(3)

gamma = WorksheetFunction.Radians(gamma)

VRotArray(1) = Cos(gamma) * X - Sin(gamma) * Y
VRotArray(2) = Sin(gamma) * X + Cos(gamma) * Y
VRotArray(3) = Z

If Index > 0 Then
    Select Case Index
        Case 1
            VRotXY = VRotArray(1)
        Case 2
            VRotXY = VRotArray(2)
        Case 3
            VRotXY = VRotArray(3)
    End Select
Else
    VRotXY = VRotArray
End If

End Function

Function IsVertical(VectorList As Variant) As Integer

Dim X As Double
Dim Y As Double
Dim Z As Double

X = VectorList(1)
Y = VectorList(2)
Z = VectorList(3)

IsVertical = 0

If X = 0 And Y = 0 Then IsVertical = 1


End Function

Function MultiplyV(VectorList As Variant, Scalar As Double) As Variant

Dim X As Double
Dim Y As Double
Dim Z As Double

Dim MultipliedV(1 To 3) As Double

X = VectorList(1)
Y = VectorList(2)
Z = VectorList(3)

MultipliedV(1) = X * Scalar
MultipliedV(2) = Y * Scalar
MultipliedV(3) = Z * Scalar

MultiplyV = MultipliedV

End Function

Function DivideV(VectorList As Variant, Scalar As Double) As Variant

Dim X As Double
Dim Y As Double
Dim Z As Double

Dim DividedV(1 To 3) As Double

X = VectorList(1)
Y = VectorList(2)
Z = VectorList(3)

DividedV(1) = X / Scalar
DividedV(2) = Y / Scalar
DividedV(3) = Z / Scalar

DivideV = DividedV

End Function
Function SubtractVs(V1 As Variant, V2 As Variant) As Variant

Dim X1 As Double
Dim Y1 As Double
Dim Z1 As Double

Dim X2 As Double
Dim Y2 As Double
Dim Z2 As Double

Dim newV(1 To 3) As Double

X1 = V1(1)
Y1 = V1(2)
Z1 = V1(3)

X2 = V2(1)
Y2 = V2(2)
Z2 = V2(3)

newV(1) = X1 - X2
newV(2) = Y1 - Y2
newV(3) = Z1 - Z2

SubtractVs = newV

End Function
Function AddVs(V1 As Variant, V2 As Variant) As Variant

Dim X1 As Double
Dim Y1 As Double
Dim Z1 As Double

Dim X2 As Double
Dim Y2 As Double
Dim Z2 As Double

Dim newV(1 To 3) As Double

X1 = V1(1)
Y1 = V1(2)
Z1 = V1(3)

X2 = V2(1)
Y2 = V2(2)
Z2 = V2(3)

newV(1) = X1 + X2
newV(2) = Y1 + Y2
newV(3) = Z1 + Z2

AddVs = newV

End Function
Function MultiplyVs(V1 As Variant, V2 As Variant) As Variant

Dim X1 As Double
Dim Y1 As Double
Dim Z1 As Double

Dim X2 As Double
Dim Y2 As Double
Dim Z2 As Double

Dim newV(1 To 3) As Double

X1 = V1(1)
Y1 = V1(2)
Z1 = V1(3)

X2 = V2(1)
Y2 = V2(2)
Z2 = V2(3)

newV(1) = X1 * X2
newV(2) = Y1 * Y2
newV(3) = Z1 * Z2

MultVs = newV

End Function

Function Rodrig(VectorList As Variant, Axis As Variant, gamma As Double, Optional Index As Integer) As Variant


Dim V1 As Double
Dim V2 As Double
Dim V3 As Double

Dim A1 As Double
Dim A2 As Double
Dim A3 As Double

Dim DotAB As Double
Dim DotBB As Double

Dim ABpar As Variant
Dim ABperp As Variant
Dim w As Variant
Dim LenABperp As Double
Dim Lenw As Double
Dim rotABperp As Variant
Dim rotA As Variant

gamma = WorksheetFunction.Radians(gamma)

'DotAB = DotProd(VectorList, Axis)
'DotBB = DotProd(Axis, Axis)

'ABpar = MultiplyV(Axis, (DotAB / DotBB))
'ABperp = SubtractVs(VectorList, ABpar)
ABperp = VectorList
w = XProd(Axis, ABperp)
LenABperp = VLen(ABperp)
Lenw = VLen(w)
rotABperp = MultiplyV(AddVs(MultiplyV(ABperp, (Cos(gamma) / LenABperp)), MultiplyV(w, (Sin(gamma) / Lenw))), LenABperp)
'rotA = rotABperp + ABpar

If Index > 0 Then
    Select Case Index
        Case 1
            Rodrig = rotABperp(1)
        Case 2
            Rodrig = rotABperp(2)
        Case 3
            Rodrig = rotABperp(3)
    End Select
Else
    Rodrig = rotABperp
End If

End Function
