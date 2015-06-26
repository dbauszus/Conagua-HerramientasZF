Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class Label

    Public Sub Label()
        Try
            Loader.SIS.CreateListFromSelection("lTXT")
            Loader.SIS.OpenSel(0)
            Dim scale = Loader.SIS.GetFlt(SIS_OT_DATASET, Loader.SIS.GetDataset(), "_scale#")
            If SetDatasetOverlay() = False Then Exit Sub
            Dim FT As String = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "_featureTable$")
            Dim FC As Integer = Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "_FC&")
            Dim seccion As Double = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "seccion#")
            Dim tramo As String = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "tramo$")
            Dim x0, x1, x2, x3, x4, y0, y1, y2, y3, y4, z As Double
            x0 = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
            y0 = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
            Try
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC& = 30")
                Loader.SIS.Snap2D(x0, y0, 3000 / scale, True, "L", "fProperty", "")
                Loader.SIS.SplitPos(x4, y4, z, (Loader.SIS.GetGeomPt(0, Loader.SIS.GetGeomNumPt(0) - 1)))
                Loader.SIS.DeselectAll()
                Loader.SIS.SelectItem()
                Loader.SIS.CreateListFromSelection("lDelete")
                Dim lResponse As Integer
                Do
                    lResponse = Loader.SIS.GetPosEx(x3, y3, z)
                    Select Case lResponse
                        Case SIS_ARG_ESCAPE
                            Exit Do
                        Case SIS_ARG_POSITION
                            Loader.SIS.Delete("lDelete")
                            Loader.SIS.MoveList("lTXT", x3 - x0, y3 - y0, 0, 0, 1)
                            Loader.SIS.SelectList("lTXT")
                            Dim plano = False
                            Try
                                Loader.SIS.DoCommand("AComTextToBox")
                            Catch
                                plano = True
                            End Try
                            Loader.SIS.SplitExtent(x1, y1, z, x2, y2, z, Loader.SIS.GetListExtent("lTXT"))
                            If (x3 - x4) < 0 Then
                                If (y3 - y4) < 0 Then
                                    Loader.SIS.MoveTo(x1, y2 + (scale / 1500), 0)
                                    Loader.SIS.LineTo(x2, y2 + (scale / 1500), 0)
                                Else
                                    Loader.SIS.MoveTo(x1, y1 - (scale / 1500), 0)
                                    Loader.SIS.LineTo(x2, y1 - (scale / 1500), 0)
                                End If
                            Else
                                If (y3 - y4) < 0 Then
                                    Loader.SIS.MoveTo(x2, y2 + (scale / 1500), 0)
                                    Loader.SIS.LineTo(x1, y2 + (scale / 1500), 0)
                                Else
                                    Loader.SIS.MoveTo(x2, y1 - (scale / 1500), 0)
                                    Loader.SIS.LineTo(x1, y1 - (scale / 1500), 0)
                                End If
                            End If
                            Loader.SIS.LineTo(x4, y4, 0)
                            Loader.SIS.UpdateItem()
                            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "seccion#", seccion)
                            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "tramo$", tramo$)
                            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", FT)
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", FC)
                            Loader.SIS.CloseItem()
                            Loader.SIS.DeselectAll()
                            Loader.SIS.SelectList("lTXT")
                            If plano = False Then Loader.SIS.DoCommand("AComBoxToText")
                            Exit Do
                    End Select
                Loop
            Catch ex As Exception
                Dim lResponse As Integer
                Do
                    lResponse = Loader.SIS.GetPosEx(x3, y3, z)
                    Select Case lResponse
                        Case SIS_ARG_ESCAPE
                            Exit Do
                        Case SIS_ARG_POSITION
                            Loader.SIS.MoveList("lTXT", x3 - x0, y3 - y0, 0, 0, 1)
                            Dim plano = False
                            Try
                                Loader.SIS.DoCommand("AComTextToBox")
                            Catch
                                plano = True
                            End Try
                            Loader.SIS.SplitExtent(x1, y1, z, x2, y2, z, Loader.SIS.GetListExtent("lTXT"))
                            If (x3 - x0) < 0 Then
                                If (y3 - y0) < 0 Then
                                    Loader.SIS.MoveTo(x1, y2 + (scale / 1500), 0)
                                    Loader.SIS.LineTo(x2, y2 + (scale / 1500), 0)
                                Else
                                    Loader.SIS.MoveTo(x1, y1 - (scale / 1500), 0)
                                    Loader.SIS.LineTo(x2, y1 - (scale / 1500), 0)
                                End If
                            Else
                                If (y3 - y0) < 0 Then
                                    Loader.SIS.MoveTo(x2, y2 + (scale / 1500), 0)
                                    Loader.SIS.LineTo(x1, y2 + (scale / 1500), 0)
                                Else
                                    Loader.SIS.MoveTo(x2, y1 - (scale / 1500), 0)
                                    Loader.SIS.LineTo(x1, y1 - (scale / 1500), 0)
                                End If
                            End If
                            Loader.SIS.LineTo(x0, y0, 0)
                            Loader.SIS.UpdateItem()
                            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "seccion#", seccion)
                            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "tramo$", tramo$)
                            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", FT)
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", FC)
                            Loader.SIS.CloseItem()
                            Loader.SIS.DeselectAll()
                            Loader.SIS.SelectList("lTXT")
                            If plano = False Then Loader.SIS.DoCommand("AComBoxToText")
                            Exit Do
                    End Select
                Loop
            End Try
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