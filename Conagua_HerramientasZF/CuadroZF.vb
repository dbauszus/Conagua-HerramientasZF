Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class CuadroZF

    Private seccion_pri As Double
    Private seccion_ult As Double
    Private tramo As String
    Private scale As Double
    Private Calc As New Calc

    Public Sub CrearCuadroZF()
        Try
            Loader.SIS.CreateListFromSelection("lSelect")
            Loader.SIS.OpenList("lSelect", 0)
            scale = Loader.SIS.GetFlt(SIS_OT_DATASET, Loader.SIS.GetDataset(), "_scale#")
            Try
                tramo = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "tramo$")
            Catch
                MsgBox("No tiene un propiedad de tramo$")
                Exit Sub
            End Try
            If SetDatasetOverlay() = False Then Exit Sub
            Loader.SIS.OpenListCursor("cursor", "lSelect", "seccion#")
            Loader.SIS.OpenSortedCursor("scursor", "cursor", 0, True)
            seccion_pri = Loader.SIS.GetCursorFieldValue("scursor", 0)
            Loader.SIS.MoveCursor("scursor", 1)
            seccion_ult = Loader.SIS.GetCursorFieldValue("scursor", 0)
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC& = 42")
            Loader.SIS.Scan("list1", "E", "fProperty", "")
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC& = 21")
            Loader.SIS.Scan("list2", "E", "fProperty", "")
            Loader.SIS.CombineLists("lVertices", "list1", "list2", SIS_BOOLEAN_OR)
            Loader.SIS.CreatePropertyFilter("fProperty", "seccion# <= " & seccion_ult.ToString)
            Loader.SIS.CreatePropertyFilter("fCombine", "seccion# >= " & seccion_pri.ToString)
            Loader.SIS.CombineFilter("fProperty", "fProperty", "fCombine", SIS_BOOLEAN_AND)
            If Loader.SIS.ScanList("lVertices", "lVertices", "fProperty", "") > 0 Then
                SetPointIndexIzq()
                SetPointIndexDer()
                Dim areaIzq As Double = CreatePoligono("I")
                Dim areaDer As Double = CreatePoligono("D")
                SetPointProperties("I")
                SetPointProperties("D")
                CreateTable("I", areaIzq)
                CreateTable("D", areaDer)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub SetPointIndexDer()
        Try
            'start at ZF
            Loader.SIS.CreatePropertyFilter("fProperty", "lado$ = 'ZD'")
            Loader.SIS.ScanList("lZF", "lVertices", "fProperty", "")
            Loader.SIS.OpenListCursor("cursor", "lZF", "seccion#")
            Loader.SIS.OpenSortedCursor("sortedCursor", "cursor", 0, True)
            Dim idx As Integer = 0
            Do
                Loader.SIS.OpenCursorItem("sortedCursor")
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "index&", idx)
                Loader.SIS.UpdateCursorItem("sortedCursor")
                idx += 1
            Loop Until Loader.SIS.MoveCursor("sortedCursor", 1) = 0

            'then return on NAMO
            Loader.SIS.CreatePropertyFilter("fProperty", "lado$ = 'ND'")
            Loader.SIS.ScanList("lNAMO", "lVertices", "fProperty", "")
            Loader.SIS.OpenListCursor("cursor", "lNAMO", "seccion#")
            Loader.SIS.OpenSortedCursor("sortedCursor", "cursor", 0, False)
            Do
                Loader.SIS.OpenCursorItem("sortedCursor")
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "index&", idx)
                Loader.SIS.UpdateCursorItem("sortedCursor")
                idx += 1
            Loop Until Loader.SIS.MoveCursor("sortedCursor", 1) = 0

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub SetPointIndexIzq()
        Try
            'start at NAMO
            Loader.SIS.CreatePropertyFilter("fProperty", "lado$ = 'NI'")
            Loader.SIS.ScanList("lNAMO", "lVertices", "fProperty", "")
            Loader.SIS.OpenListCursor("cursor", "lNAMO", "seccion#")
            Loader.SIS.OpenSortedCursor("sortedCursor", "cursor", 0, True)
            Dim idx As Integer = 0
            Do
                Loader.SIS.OpenCursorItem("sortedCursor")
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "index&", idx)
                Loader.SIS.UpdateCursorItem("sortedCursor")
                idx += 1
            Loop Until Loader.SIS.MoveCursor("sortedCursor", 1) = 0

            'then return on ZF
            Loader.SIS.CreatePropertyFilter("fProperty", "lado$ = 'ZI'")
            Loader.SIS.ScanList("lZF", "lVertices", "fProperty", "")
            Loader.SIS.OpenListCursor("cursor", "lZF", "seccion#")
            Loader.SIS.OpenSortedCursor("sortedCursor", "cursor", 0, False)
            Do
                Loader.SIS.OpenCursorItem("sortedCursor")
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "index&", idx)
                Loader.SIS.UpdateCursorItem("sortedCursor")
                idx += 1
            Loop Until Loader.SIS.MoveCursor("sortedCursor", 1) = 0

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Function CreatePoligono(ByVal lado As String)
        Try
            Loader.SIS.CreatePropertyFilter("fProperty", "lado$ = 'Z" & lado & "'")
            Loader.SIS.ScanList("lZF", "lVertices", "fProperty", "")
            Loader.SIS.CreatePropertyFilter("fProperty", "lado$ = 'N" & lado & "'")
            Loader.SIS.ScanList("lNAMO", "lVertices", "fProperty", "")
            Loader.SIS.CombineLists("list", "lZF", "lNAMO", SIS_BOOLEAN_OR)
            Loader.SIS.OpenListCursor("cursor", "list", "index&" & vbTab & "_ox#" & vbTab & "_oy#" & vbTab & "seccion#")
            Loader.SIS.OpenSortedCursor("sortedCursor", "cursor", 0, True)

            Dim xStart = Loader.SIS.GetCursorFieldValue("sortedCursor", 1)
            Dim yStart = Loader.SIS.GetCursorFieldValue("sortedCursor", 2)
            Loader.SIS.MoveTo(xStart, yStart, 0)
            Do While Loader.SIS.MoveCursor("sortedCursor", 1) = 1
                Loader.SIS.LineTo(Loader.SIS.GetCursorFieldValue("sortedCursor", 1), Loader.SIS.GetCursorFieldValue("sortedCursor", 2), 0)
            Loop
            Loader.SIS.LineTo(xStart, yStart, 0)

            'build poly and set FC
            Loader.SIS.UpdateItem()
            Loader.SIS.DeselectAll()
            Loader.SIS.SelectItem()
            Loader.SIS.DoCommand("AComFillGeometry")
            Loader.SIS.OpenSel(0)
            Dim area As Double = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_area#")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "tramo$", tramo)
            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "seccion_pri#", seccion_pri)
            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "seccion_ult#", seccion_ult)
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "lado$", "S" & lado)
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_ZF")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 40)
            Loader.SIS.CloseItem()
            Return area

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return 0
        End Try
    End Function

    Private Class V
        Public x As Double
        Public y As Double
        Public pv As String
    End Class

    Private Sub SetPointProperties(ByVal lado As String)
        Try
            Loader.SIS.CreatePropertyFilter("fProperty", "lado$ = 'Z" & lado & "'")
            Loader.SIS.ScanList("lZF", "lVertices", "fProperty", "")
            Loader.SIS.CreatePropertyFilter("fProperty", "lado$ = 'N" & lado & "'")
            Loader.SIS.ScanList("lNAMO", "lVertices", "fProperty", "")
            Loader.SIS.CombineLists("list", "lZF", "lNAMO", SIS_BOOLEAN_OR)
            Loader.SIS.OpenListCursor("cursor", "list", "index&" & vbTab & "_ox#" & vbTab & "_oy#" & vbTab & "pv$")
            Loader.SIS.OpenSortedCursor("sortedCursor", "cursor", 0, True)

            Dim distancia As Double = 0
            Dim angle As Double = 0
            Dim vStart = New V
            vStart.x = Loader.SIS.GetCursorFieldValue("sortedCursor", 1)
            vStart.y = Loader.SIS.GetCursorFieldValue("sortedCursor", 2)
            vStart.pv = Loader.SIS.GetCursorFieldValue("sortedCursor", 3)
            Dim vEst = New V
            Dim vPV = New V
            vEst.x = vStart.x
            vEst.y = vStart.y
            vEst.pv = vStart.pv
            Loader.SIS.MoveCursor("sortedCursor", 1)

            Do
                vPV.x = Loader.SIS.GetCursorFieldValue("sortedCursor", 1)
                vPV.y = Loader.SIS.GetCursorFieldValue("sortedCursor", 2)
                vPV.pv = Loader.SIS.GetCursorFieldValue("sortedCursor", 3)
                Loader.SIS.OpenCursorItem("sortedCursor")
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "est$", vEst.pv)
                distancia = Calc.Distance(vEst.x, vEst.y, 0, vPV.x, vPV.y, 0)
                Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "distancia#", String.Format("{0:###,###,##0.0000}", distancia))
                angle = Calc.GetAngle(vEst.x, vEst.y, vPV.x, vPV.y)
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "rumbo$", Calc.GetRumbo(angle))
                Loader.SIS.UpdateCursorItem("sortedCursor")
                vEst.x = vPV.x
                vEst.y = vPV.y
                vEst.pv = vPV.pv
            Loop Until Loader.SIS.MoveCursor("sortedCursor", 1) = 0

            Loader.SIS.OpenSortedCursor("sortedCursor", "cursor", 0, True)
            Loader.SIS.OpenCursorItem("sortedCursor")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "est$", vEst.pv)
            distancia = Calc.Distance(vEst.x, vEst.y, 0, vStart.x, vStart.y, 0)
            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "distancia#", String.Format("{0:###,###,##0.0000}", distancia))
            angle = Calc.GetAngle(vEst.x, vEst.y, vStart.x, vStart.y)
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "rumbo$", Calc.GetRumbo(angle))
            Loader.SIS.UpdateCursorItem("sortedCursor")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub CreateTable(ByVal lado As String, ByVal area As Double)
        Try
            Dim cadeDesde As String = Calc.FormatoCadenamiento(CDbl(seccion_pri))
            Dim cadeHasta As String = Calc.FormatoCadenamiento(CDbl(seccion_ult))
            Dim margen As String
            If lado = "I" Then margen = "izquierda" Else margen = "derecha"
            Dim tabla = String.Format("{0} - {1} del km {2} al km {3}", tramo, margen, cadeDesde, cadeHasta)

            TableOverlay(tramo)
            EmptyList("lTable")
            EmptyList("lTexts")
            EmptyList("lTextsL")
            EmptyList("lLines")

            Dim xC As Double = 20
            Dim yR As Double = 5
            Dim xTotal = xC * 8.75

            'title
            Loader.SIS.CreateText(xTotal * 0.5, -yR * 1, 0, String.Format("Cuadro de construcción de zona federal margen {0}", margen))
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 512)
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "tabla$", tabla)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'KM secc
            Loader.SIS.CreateText(xTotal * 0.5, -yR * 2, 0, String.Format("del km {0} al km {1}", cadeDesde, cadeHasta))
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 512)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'Est.
            Loader.SIS.CreateText(xC * 0.5, -yR * 4, 0, "Est.")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 512)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 0)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'P.V.
            Loader.SIS.CreateText(xC * 1.5, -yR * 4, 0, "P.V.")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 512)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 1)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'Rumbo
            Loader.SIS.CreateText(xC * 2.75, -yR * 4, 0, "Rumbo")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 512)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 2)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'Distancia (m)
            Loader.SIS.CreateText(xC * 4.125, -yR * 4, 0, "Distancia (m)")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 512)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 3)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'Vértice
            Loader.SIS.CreateText(xC * 5.25, -yR * 4, 0, "Vértice")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 512)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 4)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'Coordenadas
            Loader.SIS.CreateText(xC * 7.25, -yR * 3.5, 0, "Coordenadas")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 512)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'X
            Loader.SIS.CreateText(xC * 6.5, -yR * 4.5, 0, "X")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 512)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 5)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'Y
            Loader.SIS.CreateText(xC * 8.0, -yR * 4.5, 0, "Y")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 512)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 6)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTexts")

            'line row 3
            Loader.SIS.MoveTo(0, -yR * 3, 0)
            Loader.SIS.LineTo(xTotal, -yR * 3, 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lLines")

            'line row 4 coordenadas
            Loader.SIS.MoveTo(xC * 5.75, -yR * 4, 0)
            Loader.SIS.LineTo(xTotal, -yR * 4, 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lLines")

            Loader.SIS.CreatePropertyFilter("fProperty", "lado$ = 'Z" & lado & "'")
            Loader.SIS.ScanList("lZF", "lVertices", "fProperty", "")
            Loader.SIS.CreatePropertyFilter("fProperty", "lado$ = 'N" & lado & "'")
            Loader.SIS.ScanList("lNAMO", "lVertices", "fProperty", "")
            Loader.SIS.CombineLists("list", "lZF", "lNAMO", SIS_BOOLEAN_OR)
            Loader.SIS.OpenListCursor("cursor", "list", "index&" & vbTab & "est$" & vbTab & "pv$" & vbTab & "rumbo$" & vbTab & "distancia#" & vbTab & "_ox#" & vbTab & "_oy#")
            Loader.SIS.OpenSortedCursor("sortedCursor", "cursor", 0, True)

            Dim xt, yt, dt As Double
            Dim row = 5
            Dim firstRow = True
            Dim lastRow = False

            Do
                Loader.SIS.MoveTo(0, -yR * row, 0)
                Loader.SIS.LineTo(xTotal, -yR * row, 0)
                Loader.SIS.UpdateItem()
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 4)
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lLines")

                'Est.
                If Not firstRow = True Then
                    Loader.SIS.CreateText(xC * 0.5, -yR * (row + 0.5), 0, Loader.SIS.GetCursorFieldValue("sortedCursor", 1).ToString)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 0)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 4)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.AddToList("lTexts")
                End If

                'P.V.
                If Not firstRow = True Then
                    Loader.SIS.CreateText(xC * 1.5, -yR * (row + 0.5), 0, Loader.SIS.GetCursorFieldValue("sortedCursor", 2).ToString)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 1)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 4)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.AddToList("lTexts")
                End If

                'Rumbo
                If Not firstRow = True Then
                    Loader.SIS.CreateText(xC * 2.75, -yR * (row + 0.5), 0, Loader.SIS.GetCursorFieldValue("sortedCursor", 3).ToString)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 2)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 4)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.AddToList("lTexts")
                End If

                'Distancia (m)
                If Not firstRow = True Then
                    dt = Loader.SIS.GetCursorFieldValue("sortedCursor", 4).ToString
                    Loader.SIS.CreateText(xC * 4.125, -yR * (row + 0.5), 0, String.Format("{0:###,###,##0.0000}", dt))
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 3)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 4)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.AddToList("lTexts")
                End If

                'Vértice
                Loader.SIS.CreateText(xC * 5.25, -yR * (row + 0.5), 0, Loader.SIS.GetCursorFieldValue("sortedCursor", 2).ToString)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 4)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 4)
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lTexts")

                'X
                xt = Loader.SIS.GetCursorFieldValue("sortedCursor", 5).ToString
                Loader.SIS.CreateText(xC * 6.5, -yR * (row + 0.5), 0, String.Format("{0:##,###,###.0000}", xt))
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 5)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 4)
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lTexts")

                'Y
                yt = Loader.SIS.GetCursorFieldValue("sortedCursor", 6).ToString
                Loader.SIS.CreateText(xC * 8, -yR * (row + 0.5), 0, String.Format("{0:##,###,###.0000}", yt))
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "col&", 6)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 4)
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lTexts")

                row += 1

                If lastRow = True Then
                    Exit Do
                ElseIf firstRow = True Then
                    firstRow = False
                    Loader.SIS.MoveCursor("sortedCursor", 1)
                ElseIf Loader.SIS.MoveCursor("sortedCursor", 1) = 0 Then
                    Loader.SIS.OpenSortedCursor("sortedCursor", "cursor", 0, True)
                    lastRow = True
                End If

            Loop

            'another line
            Loader.SIS.MoveTo(0, -yR * row, 0)
            Loader.SIS.LineTo(xTotal, -yR * row, 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 4)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lLines")

            'Área
            Loader.SIS.CreateText(xC * 0.25, -yR * (row + 0.5), 0, String.Format("Superficie: {0:###,###,###.0000} m²", area))
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 509)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "fil&", row - 4)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTextsL")

            row += 1

            'draw border
            Loader.SIS.CreateRectangle(0, 0, xTotal, -yR * row)
            Loader.SIS.AddToList("lLines")

            'line column 1 
            Loader.SIS.MoveTo(xC * 1, -yR * 3, 0)
            Loader.SIS.LineTo(xC * 1, -yR * row + yR, 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lLines")

            'line column 2
            Loader.SIS.MoveTo(xC * 2, -yR * 3, 0) '2
            Loader.SIS.LineTo(xC * 2, -yR * row + yR, 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lLines")

            'line column 4
            Loader.SIS.MoveTo(xC * 3.5, -yR * 3, 0) '4
            Loader.SIS.LineTo(xC * 3.5, -yR * row + yR, 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lLines")

            'line column 6
            Loader.SIS.MoveTo(xC * 4.75, -yR * 3, 0) '6
            Loader.SIS.LineTo(xC * 4.75, -yR * row + yR, 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lLines")

            'line column 7
            Loader.SIS.MoveTo(xC * 5.75, -yR * 3, 0) '7
            Loader.SIS.LineTo(xC * 5.75, -yR * row + yR, 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lLines")

            'line column 9 (shorter)
            Loader.SIS.MoveTo(xC * 7.25, -yR * 4, 0) '9
            Loader.SIS.LineTo(xC * 7.25, -yR * row + yR, 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lLines")

            'change texts (centred)
            Loader.SIS.SetListStr("lTexts", "_featureTable$", "CONAGUA_TABLA")
            Loader.SIS.SetListInt("lTexts", "_text_alignV&", SIS_MIDDLE)
            Loader.SIS.SetListInt("lTexts", "_text_alignH&", SIS_CENTRE)

            'change texts (leftbound)
            Loader.SIS.SetListStr("lTextsL", "_featureTable$", "CONAGUA_TABLA")
            Loader.SIS.SetListInt("lTextsL", "_text_alignV&", SIS_MIDDLE)
            Loader.SIS.SetListInt("lTextsL", "_text_alignH&", SIS_LEFT)

            'change lines
            Loader.SIS.SetListStr("lLines", "_featureTable$", "CONAGUA_TABLA")
            Loader.SIS.SetListInt("lLines", "_FC&", 430)
            Loader.SIS.CombineLists("lTexts", "lTexts", "lTextsL", SIS_BOOLEAN_OR)
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
                        Loader.SIS.MoveList("lTable", xx, yy, 0, 0, scale / 1000)
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
                        Loader.SIS.SetStr(SIS_OT_DATASET, Loader.SIS.GetInt(SIS_OT_OVERLAY, i, "_nDataset&"), "_featureTable$", "CONAGUA_ZF")
                    Catch ex As Exception
                        Exit For
                    End Try
                    Exit Sub
                End If
            Next
            Loader.SIS.CreateInternalOverlay(overlayName, 0)
            Loader.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", 0)
            Loader.SIS.SetFlt(SIS_OT_DATASET, Loader.SIS.GetInt(SIS_OT_OVERLAY, 0, "_nDataset&"), "_scale#", scale)
            Loader.SIS.SetStr(SIS_OT_DATASET, Loader.SIS.GetInt(SIS_OT_OVERLAY, 0, "_nDataset&"), "_featureTable$", "CONAGUA_ZF")
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
                    Try
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "tramo$", tramo)
                    Catch
                        MsgBox("Overlay locked")
                        Return False
                    End Try
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.ToString)
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
