Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class CuadroZF_Area

    Private tramo As String
    Private planoID As String
    Private scaleFactor As Double
    Private Calc As New Calc

    Public Sub CrearCuadroZF_Area()
        Try
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC& = 808")
            If Not Loader.SIS.Scan("lPlanos", "V", "fProperty", "") > 0 Then
                MsgBox("No Marcos")
                Exit Sub
            End If
            tramo = Loader.SIS.GetListItemStr("lPlanos", 0, "tramo$")
            scaleFactor = Loader.SIS.GetListItemFlt("lPlanos", 0, "scaleFactor#")
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC& = 40")
            If Not Loader.SIS.Scan("lAreas", "V", "fProperty", "") > 0 Then
                MsgBox("No Areas")
                Exit Sub
            End If
            For i = 0 To Loader.SIS.GetListSize("lPlanos") - 1
                Loader.SIS.OpenList("lPlanos", i)
                Dim plano = String.Format("{0} de {1}", Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "planoID$"), Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "planoNUMtotal&"))
                Dim planoNUMsuffix = ""
                Try
                    planoNUMsuffix = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "planoNUMsuffix$")
                Catch ex As Exception
                End Try
                Dim planoNUM = String.Format("{0:00}{1}", Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "planoNUM&"), planoNUMsuffix)
                Loader.SIS.CreateLocusFromItem("locusContain", SIS_GT_CONTAIN, SIS_GM_GEOMETRY)
                If Loader.SIS.ScanList("lScan", "lAreas", "fProperty", "locusContain") > 0 Then
                    For ii = 0 To Loader.SIS.GetListSize("lScan") - 1
                        Loader.SIS.OpenList("lScan", ii)
                        Try
                            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "plano$", plano)
                            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "planoNUM$", planoNUM)
                        Catch ex As Exception
                            MsgBox("Overlay locked")
                            Exit Sub
                        End Try               
                        Loader.SIS.UpdateItem()
                    Next ii
                End If
            Next i

            CreateTable()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Class c_row
        Public plano As String
        Public m_izq As Double
        Public m_der As Double
    End Class

    Private Sub CreateTable()
        Try
            TableOverlay(tramo)
            EmptyList("lTable")
            EmptyList("lTexts")
            EmptyList("lLines")

            Dim tabla = tramo + " - superficie, " + planoID
            Dim xC As Double = 40
            Dim yR As Double = 5
            Dim xTotal = xC * 4.375

            'title
            Loader.SIS.CreateText(xTotal * 0.5, -yR * 1, 0, "Superficie de la zona federal (m²)")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 512)
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "tabla$", tabla)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'Plano
            Loader.SIS.CreateText(xC * 0.6, -yR * 3, 0, "Plano")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 512)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 0)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'Margen izquierda
            Loader.SIS.CreateText(xC * 1.7, -yR * 3, 0, "Margen izquierda")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 512)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 1)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'Margen derecha
            Loader.SIS.CreateText(xC * 2.7, -yR * 3, 0, "Margen derecha")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 512)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 2)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'Acumulado subtotal de ambas márgenes
            Loader.SIS.CreateText(xC * 3.7875, -yR * 3, 0, "Acumulado subtotal de" + Environment.NewLine + "ambas márgenes")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 512)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 3)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'line row 2
            Loader.SIS.MoveTo(0, -yR * 2, 0)
            Loader.SIS.LineTo(xTotal, -yR * 2, 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lLines")

            Loader.SIS.OpenListCursor("cursor", "lAreas", "planoNUM$" & vbTab & "lado$" & vbTab & "plano$" & vbTab & "_area#")
            Loader.SIS.OpenSortedCursor("sortedCursor", "cursor", 0, False)

            Dim cursorValues = Loader.SIS.GetCursorValues("cursor", 2, 0, ",", "")
            Dim uniqueCursorValues As String() = cursorValues.Split(New Char() {","c}).Distinct.ToArray
            Dim a_row(uniqueCursorValues.Length - 1) As c_row
            For i = 0 To uniqueCursorValues.Length - 1
                a_row(i) = New c_row
            Next i
            Loader.SIS.OpenSortedCursor("sortedCursor", "cursor", 0, True)
            Dim idx As Integer = 0
            For i = 0 To Loader.SIS.GetListSize("lAreas") - 1
                a_row(idx).plano = Loader.SIS.GetCursorFieldValue("sortedCursor", 2)
                If Loader.SIS.GetCursorFieldValue("sortedCursor", 1) = "SI" Then a_row(idx).m_izq += Loader.SIS.GetCursorFieldValue("sortedCursor", 3)
                If Loader.SIS.GetCursorFieldValue("sortedCursor", 1) = "SD" Then a_row(idx).m_der += Loader.SIS.GetCursorFieldValue("sortedCursor", 3)
                Loader.SIS.MoveCursor("sortedCursor", 1)
                If Not a_row(idx).plano = Loader.SIS.GetCursorFieldValue("sortedCursor", 2) Then idx += 1
            Next i

            Dim row = 4
            For i = 0 To a_row.Length - 1
                Loader.SIS.MoveTo(0, -yR * row, 0)
                Loader.SIS.LineTo(xTotal, -yR * row, 0)
                Loader.SIS.UpdateItem()
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 3)
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lLines")

                'Plano
                Loader.SIS.CreateText(xC * 0.6, -yR * (row + 0.5), 0, a_row(i).plano)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 0)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 3)
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lTexts")

                'Margen izquierda
                Loader.SIS.CreateText(xC * 1.7, -yR * (row + 0.5), 0, String.Format("{0:#.####}", a_row(i).m_izq))
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 1)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 3)
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lTexts")

                'Margen derecha
                Loader.SIS.CreateText(xC * 2.7, -yR * (row + 0.5), 0, String.Format("{0:#.####}", a_row(i).m_der))
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 2)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 3)
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lTexts")

                'Acumulado subtotal de ambas márgenes
                Loader.SIS.CreateText(xC * 3.7875, -yR * (row + 0.5), 0, String.Format("{0:#.####}", a_row(i).m_izq + a_row(i).m_der))
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 3)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 3)
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lTexts")

                row += 1
            Next

            'another line
            Loader.SIS.MoveTo(0, -yR * row, 0)
            Loader.SIS.LineTo(xTotal, -yR * row, 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 3)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lLines")

            'Área
            Loader.SIS.CreateText(xC * 0.6, -yR * (row + 0.5), 0, "Acumulada subtotal")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 3)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            Dim margen_total As Double

            'Margen izquierda
            margen_total = 0
            For Each r As c_row In a_row
                margen_total += r.m_izq
            Next
            Loader.SIS.CreateText(xC * 1.7, -yR * (row + 0.5), 0, String.Format("{0:#.####}", margen_total))
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 1)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 3)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'Margen izquierda
            margen_total = 0
            For Each r As c_row In a_row
                margen_total += r.m_der
            Next
            Loader.SIS.CreateText(xC * 2.7, -yR * (row + 0.5), 0, String.Format("{0:#.####}", margen_total))
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 2)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 3)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'Acumulada
            margen_total = 0
            For Each r As c_row In a_row
                margen_total += (r.m_der + r.m_izq)
            Next
            Loader.SIS.CreateText(xC * 3.7875, -yR * (row + 0.5), 0, String.Format("{0:#.####}", margen_total))
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 2)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 3)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            row += 1

            'draw border
            Loader.SIS.CreateRectangle(0, 0, xTotal, -yR * row)
            Loader.SIS.AddToList("lLines")

            'line column 1 
            Loader.SIS.MoveTo(xC * 1.2, -yR * 2, 0)
            Loader.SIS.LineTo(xC * 1.2, -yR * row, 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lLines")

            'line column 2
            Loader.SIS.MoveTo(xC * 2.2, -yR * 2, 0)
            Loader.SIS.LineTo(xC * 2.2, -yR * row, 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lLines")

            'line column 3
            Loader.SIS.MoveTo(xC * 3.2, -yR * 2, 0)
            Loader.SIS.LineTo(xC * 3.2, -yR * row, 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lLines")

            'change texts (centred)
            Loader.SIS.SetListStr("lTexts", "_featureTable$", "CONAGUA_TABLA")
            Loader.SIS.SetListInt("lTexts", "_text_alignV&", SIS_MIDDLE)
            Loader.SIS.SetListInt("lTexts", "_text_alignH&", SIS_CENTRE)

            'change lines
            Loader.SIS.SetListStr("lLines", "_featureTable$", "CONAGUA_TABLA")
            Loader.SIS.SetListInt("lLines", "_FC&", 430)
            Loader.SIS.CombineLists("lTable", "lTexts", "lLines", SIS_BOOLEAN_OR)
            Loader.SIS.SetListStr("lTable", "tabla$", tabla)

            'drop table
            Dim xx, yy, zz As Double
            Dim lResponse As Integer
            Do
                lResponse = Loader.SIS.GetPosEx(xx, yy, zz)
                Select Case lResponse
                    Case SIS_ARG_ENTER
                        Loader.SIS.Delete("lTable")
                        Exit Sub
                    Case SIS_ARG_ESCAPE
                        Loader.SIS.Delete("lTable")
                        Exit Sub
                    Case SIS_ARG_BACKSPACE
                        Loader.SIS.Delete("lTable")
                        Exit Sub
                    Case SIS_ARG_POSITION
                        Loader.SIS.MoveList("lTable", xx, yy, 0, 0, scaleFactor)
                        Loader.SIS.SetListInt("lTable", "parte&", 0)
                        Exit Do
                End Select
            Loop

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub TableOverlay(ByVal tramo As String)
        Try
            Dim overlayName = String.Format("{0}_Tabla", tramo)
            Dim overlayNameBDS = overlayName + ".bds"
            For i = 0 To Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1
                If Loader.SIS.GetStr(SIS_OT_OVERLAY, i, "_name$") = overlayName Or Loader.SIS.GetStr(SIS_OT_OVERLAY, i, "_name$") = overlayNameBDS Then
                    Loader.SIS.SetInt(SIS_OT_OVERLAY, i, "_status&", 3)
                    Loader.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", i)
                    Try
                        Loader.SIS.SetStr(SIS_OT_DATASET, Loader.SIS.GetInt(SIS_OT_OVERLAY, i, "_nDataset&"), "_featureTable$", "CONAGUA_TABLA")
                    Catch ex As Exception
                        Exit For
                    End Try
                    Exit Sub
                End If
            Next
            Loader.SIS.CreateInternalOverlay(overlayName, 0)
            Loader.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", 0)
            Loader.SIS.SetFlt(SIS_OT_DATASET, Loader.SIS.GetInt(SIS_OT_OVERLAY, 0, "_nDataset&"), "_scale#", scaleFactor * 1000)
            Loader.SIS.SetStr(SIS_OT_DATASET, Loader.SIS.GetInt(SIS_OT_OVERLAY, 0, "_nDataset&"), "_featureTable$", "CONAGUA_TABLA")
        Catch ex As Exception
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
