Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants
Imports System.Windows.Forms
Imports System.IO

Public Class Margen

    Public Sub New()
        InitializeComponent()
        Try
            cbMargen.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub btnCrearMarco_Click(sender As System.Object, e As System.EventArgs) Handles btnCrearMargen.Click
        Try
            Dim scaleFactor As Double
            Try
                scaleFactor = Loader.SIS.GetFlt(SIS_OT_DATASET, Loader.SIS.GetInt(SIS_OT_OVERLAY, Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&"), "_nDataset&"), "_scale#") / 1000
            Catch
                MsgBox("No hay overlay activo")
                Loader.SIS.Dispose()
                Loader.SIS = Nothing
                Exit Sub
            End Try
            Dim topText As String = "top"
            Dim bottomeText As String = "bottom"
            Select Case cbMargen.SelectedIndex
                Case 0
                    topText = "LÍMITE DE ZONA FEDERAL"
                    bottomeText = "MARGEN IZQUIERDA"
                Case 1
                    topText = "LÍMITE DE ZONA FEDERAL"
                    bottomeText = "MARGEN DERECHA"
                Case 2
                    topText = "LÍMITE DE CAUCE"
                    bottomeText = "(N.A.M.O.)"
            End Select
            Dim lResponse As Integer
            Dim x, y, z As Double
            Do
                lResponse = Loader.SIS.GetPosEx(x, y, z)
                Select Case lResponse
                    Case SIS_ARG_ESCAPE
                        Exit Do
                    Case SIS_ARG_POSITION
                        EmptyList("lMargen")
                        Loader.SIS.CreateText(x, y + 3 * scaleFactor, 0, topText)
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_ZF")
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 28)
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                        Loader.SIS.UpdateItem()
                        Loader.SIS.AddToList("lMargen")
                        Loader.SIS.CloseItem()
                        Loader.SIS.CreateText(x, y - 3 * scaleFactor, 0, bottomeText)
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_ZF")
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 28)
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                        Loader.SIS.UpdateItem()
                        Loader.SIS.AddToList("lMargen")
                        Loader.SIS.CloseItem()
                        Exit Do
                End Select
            Loop
            Loader.SIS.DeselectAll()
            Loader.SIS.SelectList("lMargen")
            Loader.SIS.Dispose()
            Loader.SIS = Nothing
            Me.Close()
        Catch ex As Exception
            Loader.SIS.Dispose()
            Loader.SIS = Nothing
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub EmptyList(ByVal List As String)
        Try
            Loader.SIS.EmptyList(List)
        Catch
        End Try
    End Sub

End Class