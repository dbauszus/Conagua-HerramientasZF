Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

<GisLinkProgram("Conagua_HerramientasZF")> _
Public Class Loader

    Private Shared APP As SisApplication
    Private Shared _sis As MapModeller

    Public Shared Property SIS As MapModeller
        Get
            If _sis Is Nothing Then _sis = APP.TakeoverMapManager
            Return _sis
        End Get
        Set(ByVal value As MapModeller)
            _sis = value
        End Set
    End Property

    Public Sub New(ByVal SISApplication As SisApplication)
        APP = SISApplication

        SIS.CreatePropertyFilter("ctxLabel", "_FC& = 30")
        SIS.CreatePropertyFilter("ctxTabla", "_FC& = 509")
        SIS.CreatePropertyFilter("ctxSeccion", "_FC& = 25")
        SIS.CreatePropertyFilter("btnSeccion", "_FC& = 25")
        SIS.CreatePropertyFilter("btnConagua_Plano", "_FC& = 808")
        SIS.Dispose()

        Dim group As SisRibbonGroup = APP.RibbonGroup
        group.Text = "Herramientas ZF"

        Dim ctxAcomodaSecc As SisMenuItem = New SisMenuItem("Secciones", New SisClickHandler(AddressOf subAcomodaSecc))
        ctxAcomodaSecc.Filter = "ctxSeccion"
        ctxAcomodaSecc.MinSelection = 1
        ctxAcomodaSecc.Image = My.Resources.SECCIONES
        ctxAcomodaSecc.Help = "Secciones"
        APP.ContextMenu.MenuItems.Add(ctxAcomodaSecc)

        Dim btnCuadroZF As SisRibbonButton = New SisRibbonButton("Crea el Cuadro de Construcción", New SisClickHandler(AddressOf subCuadroZF))
        btnCuadroZF.LargeImage = True
        btnCuadroZF.Filter = "btnSeccion"
        btnCuadroZF.MinSelection = 2
        btnCuadroZF.MaxSelection = 2
        btnCuadroZF.Icon = My.Resources.TABLE
        btnCuadroZF.Help = "Genera el Cuadro de Construcción de la Zona Federal en Rios"
        group.Controls.Add(btnCuadroZF)

        Dim btnCuadroZF_Area As SisRibbonButton = New SisRibbonButton("Crea el Cuadro de Area", New SisClickHandler(AddressOf subCuadroZF_Area))
        btnCuadroZF_Area.LargeImage = True
        btnCuadroZF_Area.Icon = My.Resources.TABLE
        btnCuadroZF_Area.Help = "Genera el Cuadro de Area de la Zona Federal en Rios"
        group.Controls.Add(btnCuadroZF_Area)

        Dim btnMargen As SisRibbonButton = New SisRibbonButton("Margen", New SisClickHandler(AddressOf subMargen))
        btnMargen.LargeImage = True
        btnMargen.Icon = My.Resources.MARGEN
        btnMargen.Help = "Genera Margen de Zona Federal"
        group.Controls.Add(btnMargen)

        Dim ctxLabel As SisMenuItem = New SisMenuItem("VÉRTICE ZF", New SisClickHandler(AddressOf subLabel))
        ctxLabel.Filter = "ctxLabel"
        ctxLabel.MinSelection = 1
        ctxLabel.MaxSelection = 1
        ctxLabel.Image = My.Resources.LABEL
        APP.ContextMenu.MenuItems.Add(ctxLabel)

        Dim ctxTablaSeleccion As SisMenuItem = New SisMenuItem("TABLA SELECCION", New SisClickHandler(AddressOf subTablaSeleccion))
        ctxTablaSeleccion.Filter = "ctxTabla"
        ctxTablaSeleccion.MinSelection = 1
        ctxTablaSeleccion.MaxSelection = 1
        ctxTablaSeleccion.Image = My.Resources.TABLA_SELECCION
        APP.ContextMenu.MenuItems.Add(ctxTablaSeleccion)

        Dim ctxTablaSplit As SisMenuItem = New SisMenuItem("TABLA SPLIT", New SisClickHandler(AddressOf subTablaSplit))
        ctxTablaSplit.Filter = "ctxTabla"
        ctxTablaSplit.MinSelection = 1
        ctxTablaSplit.MaxSelection = 1
        ctxTablaSplit.Image = My.Resources.TABLA_SPLIT
        APP.ContextMenu.MenuItems.Add(ctxTablaSplit)

        Dim ctxTablaClear As SisMenuItem = New SisMenuItem("TABLA CLEAR", New SisClickHandler(AddressOf subTablaClear))
        ctxTablaClear.Filter = "ctxTabla"
        ctxTablaClear.MinSelection = 1
        ctxTablaClear.MaxSelection = 1
        ctxTablaClear.Image = My.Resources.CLEAR_TABLE
        APP.ContextMenu.MenuItems.Add(ctxTablaClear)

    End Sub

    Private Sub subAcomodaSecc(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim AcomodaSecc As New AcomodaSecc
            AcomodaSecc.AcomodaSecc()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub subCuadroZF(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim CuadroZF As New CuadroZF
            CuadroZF.CrearCuadroZF()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub subCuadroZF_Area(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim CuadroZF_Area As New CuadroZF_Area
            CuadroZF_Area.CrearCuadroZF_Area()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub subMargen(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim Margen As New Margen With {.TopMost = True}
            Margen.Show()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub subLabel(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim Label As New Label
            Label.Label()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub subTablaSeleccion(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim Tabla As New Tabla
            Tabla.TablaSeleccion()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub subTablaSplit(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim Tabla As New Tabla
            Tabla.TablaSplit()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub subTablaClear(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim Tabla As New Tabla
            Tabla.TablaClear()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

End Class