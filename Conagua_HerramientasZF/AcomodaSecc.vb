Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class AcomodaSecc

    'NAMO_LIMITE (20)
    'NAMO_VERTICE (21)
    'SECCION (25)
    'TXT_ZF_CADENAMIENTO (26)
    'TXT_ZF_MARGEN (28)
    'TXT_ZF_VERTICE (30)
    'ZF_AREA (40)
    'ZF_LIMITE (41)
    'ZF_VERTICE (42)

    Private Calc As New Calc
    Private scale As Double
    Private tramo As String

    Public Sub AcomodaSecc()
        Try
            Loader.SIS.CreateListFromSelection("lSecciones")
            Loader.SIS.OpenList("lSecciones", 0)
            scale = Loader.SIS.GetFlt(SIS_OT_DATASET, Loader.SIS.GetDataset(), "_scale#")
            If SetDatasetOverlay() = False Then Exit Sub
            For iSeccion = 0 To Loader.SIS.GetListSize("lSecciones") - 1
                Loader.SIS.OpenList("lSecciones", iSeccion)
                Dim seccion As Double
                Try
                    seccion = Math.Round(Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "seccion#"), 2)
                Catch ex As Exception
                    MsgBox("Seccion invalida, falta propiedad")
                    Loader.SIS.DeselectAll()
                    Loader.SIS.SelectItem()
                    Loader.SIS.DoCommand("AComZoomSelect")
                    Exit Sub
                End Try
                Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "seccion#", seccion)
                Loader.SIS.UpdateItem()
                tramo = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "tramo$")
                Loader.SIS.CreatePropertyFilter("fTramo", "tramo$='" & tramo & "'")
                Dim LND = SetIntersectPoint(iSeccion, seccion, "LND", 21)
                If LND = 0 Then Exit Sub
                Dim LNI = SetIntersectPoint(iSeccion, seccion, "LNI", 21)
                If LNI = 0 Then Exit Sub
                Dim LZD = SetIntersectPoint(iSeccion, seccion, "LZD", 42)
                If LZD = 0 Then Exit Sub
                Dim LZI = SetIntersectPoint(iSeccion, seccion, "LZI", 42)
                If LZI = 0 Then Exit Sub

                'clean line
                Loader.SIS.OpenList("lSecciones", iSeccion)
                Loader.SIS.DeselectAll()
                Loader.SIS.SelectItem()
                Loader.SIS.CreateListFromSelection("lSelect")
                Loader.SIS.CleanLines("lSelect", 0.1, SIS_CLEAN_LINE_NONE)

                Cadenamiento(seccion, LND, LNI)

            Next iSeccion
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Function SetIntersectPoint(ByVal iSeccion As Integer, ByVal seccion As Double, ByVal lado As String, ByVal FC As Integer)
        Try
            Loader.SIS.CreatePropertyFilter("fLado", "lado$='" & lado & "'")
            Loader.SIS.CombineFilter("fIntersect", "fTramo", "fLado", SIS_BOOLEAN_AND)
            Loader.SIS.OpenList("lSecciones", iSeccion)
            If Not Loader.SIS.ScanGeometry("lIntersect", SIS_GT_INTERSECT, SIS_GM_GEOMETRY, "fIntersect", "") = 1 Then
                MsgBox("Verificar seccion, possible duplicado o " & lado & " incompleto")
                Loader.SIS.DeselectAll()
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComZoomSelect")
                Return 0
            End If
            Dim x, y, z As Double
            Dim D() As String = Split(Loader.SIS.GetGeomIntersections(0, "lIntersect"), ",")
            Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, CDbl(D(0))))
            Try
                Loader.SIS.SplitPos(x, y, z, Loader.SIS.Snap2D(x, y, 25, False, "V", "fIntersect", ""))
                Loader.SIS.OpenList("lSecciones", iSeccion)
                Loader.SIS.InsertGeomPt(0, CDbl(D(0)), x, y, 0)
                Loader.SIS.UpdateItem()

                'create point item
                Loader.SIS.CreatePoint(x, y, 0, "", 0, 1)
                Loader.SIS.UpdateItem()
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_ZF")
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", FC)
                Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "seccion#", seccion)
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "lado$", lado.Substring(1))
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "tramo$", tramo)
                Dim pv As String = lado.Substring(1) & Math.Floor(seccion).ToString
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "pv$", pv)
                Loader.SIS.CloseItem()

                'create label item
                Loader.SIS.CreateText(x, y, 0, pv)
                Loader.SIS.UpdateItem()
                Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "seccion#", seccion)
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "tramo$", tramo)
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_ZF")
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 30)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                Loader.SIS.CloseItem()

            Catch
                MsgBox("No Ve)rtice within 25m")
                Loader.SIS.DeselectAll()
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComZoomSelect")
                Return 0
            End Try
            Return CDbl(D(0))
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return 0
        End Try
    End Function

    Private Sub Cadenamiento(ByVal seccion As Double, ByVal LND As Double, ByVal LNI As Double)
        Try
            Dim x, y, z, a, aDeg As Double
            'Loader.SIS.CreatePropertyFilter("fLado", "lado$='LND'")
            'Loader.SIS.CombineFilter("fIntersect", "fTramo", "fLado", SIS_BOOLEAN_AND)
            'Loader.SIS.ScanGeometry("lIntersect", SIS_GT_INTERSECT, SIS_GM_GEOMETRY, "fIntersect", "")
            'Dim D_LND() As String = Split(Loader.SIS.GetGeomIntersections(0, "lIntersect"), ",")
            'Loader.SIS.CreatePropertyFilter("fLado", "lado$='LNI'")
            'Loader.SIS.CombineFilter("fIntersect", "fTramo", "fLado", SIS_BOOLEAN_AND)
            'Loader.SIS.ScanGeometry("lIntersect", SIS_GT_INTERSECT, SIS_GM_GEOMETRY, "fIntersect", "")
            'Dim D_LNI() As String = Split(Loader.SIS.GetGeomIntersections(0, "lIntersect"), ",")
            'Dim length As Double = ((Convert.ToDouble(D_LND(0)) + Convert.ToDouble(D_LNI(0))) / 2)
            Dim length As Double = ((LND + LNI) / 2)
            Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, length))
            a = Loader.SIS.GetGeomAngleFromLength(0, length)
            aDeg = a * (180 / Math.PI)
            If aDeg < -90 Or aDeg > 90 Then aDeg += 180
            Loader.SIS.CreateText(x + ((3 * scale / 1000) * Math.Cos(a + 1.57079)), y + ((3 * scale / 1000) * Math.Sin(a + 1.57079)), 0, Calc.FormatoCadenamiento(seccion))
            Loader.SIS.UpdateItem()
            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "seccion#", seccion)
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "tramo$", tramo)
            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", aDeg)
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_ZF")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 26)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
            Loader.SIS.CloseItem()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Function SetDatasetOverlay()
        Try
            For i = 0 To Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1
                If Loader.SIS.GetInt(SIS_OT_OVERLAY, i, "_nDataset&") = Loader.SIS.GetDataset() Then
                    Loader.SIS.SetInt(SIS_OT_OVERLAY, i, "_status&", 3)
                    Loader.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", i)
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_ox#", Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#"))
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox("Layer locked")
            Return False
        End Try
    End Function

End Class