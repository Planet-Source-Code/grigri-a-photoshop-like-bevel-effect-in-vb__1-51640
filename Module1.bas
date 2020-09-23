Attribute VB_Name = "Module1"
Option Explicit

Public Type Triplet
    X As Double
    Y As Double
    Z As Double
End Type

Public Const PI As Double = 3.14159265358979

Public Function NullTriplet() As Triplet
    '
End Function

Public Function MakeTriplet(ByVal X As Double, ByVal Y As Double, ByVal Z As Double) As Triplet
    With MakeTriplet
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function

Public Function AddTriplet(t1 As Triplet, t2 As Triplet) As Triplet
    With AddTriplet
        .X = t1.X + t2.X
        .Y = t1.Y + t2.Y
        .Z = t1.Z + t2.Z
    End With
End Function

Public Function SubTriplet(t1 As Triplet, t2 As Triplet) As Triplet
    With SubTriplet
        .X = t1.X - t2.X
        .Y = t1.Y - t2.Y
        .Z = t1.Z - t2.Z
    End With
End Function

Public Function NormTriplet(t1 As Triplet) As Double
    With t1
        NormTriplet = Sqr(.X * .X + .Y * .Y + .Z * .Z)
    End With
End Function

Public Sub NormalizeTriplet(t1 As Triplet)
    Dim tmp As Double
    With t1
        tmp = Sqr(.X * .X + .Y * .Y + .Z * .Z)
        If tmp = 0 Then Exit Sub
        tmp = 1# / tmp
        .X = .X * tmp
        .Y = .Y * tmp
        .Z = .Z * tmp
    End With
End Sub

Public Function DotTriplet(t1 As Triplet, t2 As Triplet) As Double
    DotTriplet = t1.X * t2.X + t1.Y * t2.Y + t1.Z * t2.Z
End Function

Public Function CrossTriplet(t1 As Triplet, t2 As Triplet) As Triplet
    With CrossTriplet
        .X = t1.Y * t2.Z - t2.Z * t1.Y
        .Y = -t1.X * t2.Z + t1.Z * t2.X
        .Z = t1.X * t2.Y - t1.Y * t2.X
    End With
End Function

Public Function CosAngleTriplets(t1 As Triplet, t2 As Triplet) As Double
    CosAngleTriplets = DotTriplet(t1, t2) / (NormTriplet(t1) * NormTriplet(t2))
End Function

Public Function DegreesToRadians(ByVal deg As Double) As Double
    DegreesToRadians = PI * deg / 180#
End Function

Public Function RadiansToDegrees(ByVal deg As Double) As Double
    RadiansToDegrees = 180# * deg / PI
End Function

