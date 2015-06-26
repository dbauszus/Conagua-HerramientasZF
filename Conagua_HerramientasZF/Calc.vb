Public Class Calc

    Public Function GetAngle(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Single
        Return (atan(CDbl(y2) - CDbl(y1), CDbl(x2) - CDbl(x1)) * 180 / Math.PI)
    End Function

    'Devuelve el arco tangente en radianes de dos números:
    Public Function atan(ByVal x As Double, ByVal y As Double) As Double
        Dim Theta As Double
        If (Math.Abs(x) < 0.0000001) Then
            If (Math.Abs(y) < 0.0000001) Then
                Theta = 0.0#
            ElseIf (y > 0.0#) Then
                Theta = 1.5707963267949
            Else
                Theta = -1.5707963267949
            End If
        Else
            Theta = Math.Atan(y / x)
            If (x < 0) Then
                If (y >= 0.0#) Then
                    Theta = 3.14159265358979 + Theta
                Else
                    Theta = Theta - 3.14159265358979
                End If
            End If
        End If
        Return Theta
    End Function

    Public Function Distance(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double)
        Dim X As Double = (x2 - x1) * (x2 - x1)
        Dim Y As Double = (y2 - y1) * (y2 - y1)
        Dim Z As Double = (z2 - z1) * (z2 - z1)
        Dim Dist As Double = Math.Sqrt(X + Y + Z)
        Return Dist
    End Function

    Public Function GetRumbo(ByVal angle As Double)
        Dim degrees As Double
        Dim minutes As Double
        Dim seconds As Double
        Dim sector As Double
        Dim rumbo As String = ""

        If angle > 0 Then
            If angle < 90 Then
                sector = 1
            Else
                sector = 2
            End If
        ElseIf angle < 0 Then
            angle = angle * (-1)
            angle = 360 - angle
            If angle < 270 Then
                sector = 3
            Else
                sector = 4
            End If
        End If

        'Crea Rumbos - Define el sector en la Rosa de los Vientos 1 = NE, 2 SE , 3 SW, 4 NW
        Select Case sector
            Case 1
                Degrees = Int(angle)
                Minutes = (angle - Degrees) * 60
                Seconds = (Minutes - Int(Minutes)) * 60
                Rumbo = String.Format("N{0:00}°{1:00}'{2:00.00}''E", Degrees, Int(Minutes), Seconds)
            Case 2
                angle = angle - 90
                angle = 90 - angle
                Degrees = Int(angle)
                Minutes = (angle - Degrees) * 60
                Seconds = (Minutes - Int(Minutes)) * 60
                Rumbo = String.Format("S{0:00}°{1:00}'{2:00.00}''E", Degrees, Int(Minutes), Seconds)
            Case 3
                angle = angle - 180
                Degrees = Int(angle)
                Minutes = (angle - Degrees) * 60
                Seconds = (Minutes - Int(Minutes)) * 60
                Rumbo = String.Format("S{0:00}°{1:00}'{2:00.00}''W", Degrees, Int(Minutes), Seconds)
            Case 4
                angle = angle - 270
                angle = 90 - angle
                Degrees = Int(angle)
                Minutes = (angle - Degrees) * 60
                Seconds = (Minutes - Int(Minutes)) * 60
                Rumbo = String.Format("N{0:00}°{1:00}'{2:00.00}''W", Degrees, Int(Minutes), Seconds)
        End Select

        Return Rumbo

    End Function

    Public Function FormatoCadenamiento(ByVal seccion As Double)
        Try
            Dim first = Math.Truncate((seccion / 1000))
            Dim second = Math.Round((seccion - (first * 1000)), 2)
            Dim cadenamiento = String.Format("{0}+{1:000.##}", first, second)
            Return cadenamiento
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return ""
        End Try
    End Function

End Class
