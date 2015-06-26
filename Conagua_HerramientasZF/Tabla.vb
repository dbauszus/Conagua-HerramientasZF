Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class Tabla

    Public Sub TablaSeleccion()
        Try
            Loader.SIS.OpenSel(0)
            Loader.SIS.CreatePropertyFilter("fProperty", "tabla$ = '" & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "tabla$") & "'")
            Loader.SIS.CreatePropertyFilter("fCombine", "parte& = " & Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "parte&").ToString)
            Loader.SIS.CombineFilter("fProperty", "fProperty", "fCombine", SIS_BOOLEAN_AND)
            Loader.SIS.ScanDataset("lTabla", Loader.SIS.GetDataset(), "fProperty", "")
            Loader.SIS.DeselectAll()
            Loader.SIS.SelectList("lTabla")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub TablaSplit()
        Try
            Loader.SIS.CreateClassTreeFilter("fLinea", "-Item +Line")
            Loader.SIS.CreateClassTreeFilter("fArea", "-Item +Area")

            'lTabla
            Loader.SIS.OpenSel(0)
            If SetDatasetOverlay() = False Then Exit Sub
            Dim fil = Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "fil&")
            Dim parte = Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "parte&")
            Loader.SIS.CreatePropertyFilter("fProperty", "tabla$ = '" & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "tabla$") & "'")
            Loader.SIS.CreatePropertyFilter("fCombine", "parte& = " & Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "parte&").ToString)
            Loader.SIS.CombineFilter("fProperty", "fProperty", "fCombine", SIS_BOOLEAN_AND)
            Loader.SIS.ScanDataset("lTabla", Loader.SIS.GetDataset(), "fProperty", "")

            'lSplit
            Loader.SIS.CreatePropertyFilter("fProperty", "fil& >= " & fil.ToString)
            Loader.SIS.ScanList("lSplit", "lTabla", "fProperty", "")

            'lFil
            Loader.SIS.CreatePropertyFilter("fProperty", "fil& = " & fil.ToString)
            Loader.SIS.CombineFilter("fProperty", "fProperty", "fLinea", SIS_BOOLEAN_AND)
            Loader.SIS.ScanList("lFil", "lTabla", "fProperty", "")
            Dim x0, x1, y0, y1, z As Double
            Loader.SIS.OpenList("lFil", 0)
            Loader.SIS.SplitPos(x0, y0, z, Loader.SIS.GetGeomPt(0, 0))
            Loader.SIS.SplitPos(x1, y1, z, Loader.SIS.GetGeomPt(0, 1))
            Loader.SIS.Delete("lFil")

            'lSplitArea
            Loader.SIS.ScanList("lSplitArea", "lTabla", "fArea", "")
            Loader.SIS.CopyListItems("lSplitArea")
            Loader.SIS.OpenList("lSplitArea", 0)
            Loader.SIS.SetGeomPt(0, 0, x0, y0, 0)
            Loader.SIS.SetGeomPt(0, 3, x1, y1, 0)
            Loader.SIS.CombineLists("lSplit", "lSplit", "lSplitArea", SIS_BOOLEAN_OR)

            'snap lines
            Loader.SIS.ScanList("lSnip", "lTabla", "fLinea", "")
            Loader.SIS.CopyListItems("lSnip")
            Loader.SIS.SnipGeometry("lSnip", False)
            Loader.SIS.CombineLists("lSplit", "lSplit", "lSnip", SIS_BOOLEAN_OR)

            'lArea
            Loader.SIS.ScanList("lArea", "lTabla", "fArea", "")
            Loader.SIS.OpenList("lArea", 0)
            Loader.SIS.SetGeomPt(0, 1, x0, y0, 0)
            Loader.SIS.SetGeomPt(0, 2, x1, y1, 0)

            'snap lines
            Loader.SIS.ScanList("lSnip", "lTabla", "fLinea", "")
            Loader.SIS.SnipGeometry("lSnip", False)

            'select lSplit
            Loader.SIS.DeselectAll()
            Loader.SIS.SelectList("lSplit")
            Loader.SIS.SetListInt("lSplit", "parte&", parte + 1)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub TablaClear()
        Try
            Loader.SIS.OpenSel(0)
            Loader.SIS.CreatePropertyFilter("fProperty", "tabla$ = '" & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "tabla$") & "'")
            Loader.SIS.CreatePropertyFilter("fCombine", "parte& = " & Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "parte&").ToString)
            Loader.SIS.CombineFilter("fProperty", "fProperty", "fCombine", SIS_BOOLEAN_AND)
            Loader.SIS.ScanDataset("lTabla", Loader.SIS.GetDataset(), "fProperty", "")
            Loader.SIS.CreateClassTreeFilter("fArea", "-Item +Area")
            Loader.SIS.CreateClassTreeFilter("fLine", "-Item +Line")
            Loader.SIS.CreateClassTreeFilter("fLineArea", "-Item +Line +Area")
            Loader.SIS.CreateClassTreeFilter("fBoxText", "-Item +Line +BoxText")

            'clear table
            If Loader.SIS.ScanList("lTablaArea", "lTabla", "fArea", "") = 1 Then
                Loader.SIS.OpenList("lTablaArea", 0)
                Loader.SIS.CreateLocusFromItem("locus_TablaArea", SIS_GT_INTERSECT, SIS_GM_GEOMETRY)
                Loader.SIS.ScanDataset("lIntersect", Loader.SIS.GetDataset(), "", "locus_TablaArea")
                Loader.SIS.CombineLists("lIntersect", "lIntersect", "lTabla", SIS_BOOLEAN_XOR)
                If Loader.SIS.ScanList("lSnip", "lIntersect", "fLineArea", "") > 0 Then Loader.SIS.SnipGeometry("lSnip", True)
                If Loader.SIS.ScanList("lDelete", "lIntersect", "fBoxText", "") > 0 Then Loader.SIS.Delete("lDelete")
            End If

            'change texts
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC& = 512")
            If Loader.SIS.ScanList("lText", "lTabla", "fProperty", "") > 0 Then
                Loader.SIS.OpenList("lText", 0)
                Dim scale = Loader.SIS.GetFlt(SIS_OT_DATASET, Loader.SIS.GetDataset, "_scale#")
                Loader.SIS.SelectList("lText")
                Try
                    Loader.SIS.DoCommand("AComTextToBox")
                Catch
                End Try
                Loader.SIS.SetListFlt("lText", "_character_height#", 4 * scale / 1000)
                Loader.SIS.DeselectAll()
            End If
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC& = 509")
            If Loader.SIS.ScanList("lText", "lTabla", "fProperty", "") > 0 Then
                Loader.SIS.OpenList("lText", 0)
                Dim scale = Loader.SIS.GetFlt(SIS_OT_DATASET, Loader.SIS.GetDataset, "_scale#")
                Loader.SIS.SelectList("lText")
                Try
                    Loader.SIS.DoCommand("AComTextToBox")
                Catch
                End Try
                Loader.SIS.SetListFlt("lText", "_character_height#", 2.6666666666 * scale / 1000)
                Loader.SIS.DeselectAll()
            End If

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

    Private Sub EmptyList(ByVal List As String)
        Try
            Loader.SIS.EmptyList(List)
        Catch
        End Try
    End Sub

End Class
