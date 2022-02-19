Imports System.Windows.Xps.Packaging
Imports System.IO
Imports System.IO.Packaging
Imports AGVBW

Partial Public Class fPrintDialog

    Private mp_lColumns As Integer
    Private mp_lRows As Integer
    Private mp_lPageNumber As Integer
    Private mp_lXMargin As Integer
    Private mp_lYMargin As Integer
    Private mp_oControl As ActiveGanttVBWCtl
    Private mp_dtStartDate As AGVBW.DateTime
    Private mp_dtEndDate As AGVBW.DateTime
    Private mp_lRow As Integer
    Private mp_lColumn As Integer
    Private mp_lPageWidth As Integer
    Private mp_lPageHeight As Integer
    Public m_lXPreviewMargin As Long
    Public m_lYPreviewMargin As Long
    Public sURI As String = "memorystream://Preview.xps"
    Public oURI As Uri
    Private mp_bLoaded As Boolean = False

    Public Sub New(ByRef oControl As ActiveGanttVBWCtl)
        InitializeComponent()
        mp_oControl = oControl
        mp_dtStartDate = New AGVBW.DateTime()
        mp_dtEndDate = New AGVBW.DateTime()
    End Sub

    Public Sub New(ByRef oControl As ActiveGanttVBWCtl, ByVal dtStartDate As AGVBW.DateTime, ByVal dtEndDate As AGVBW.DateTime)
        InitializeComponent()
        mp_oControl = oControl
        mp_dtStartDate = dtStartDate
        mp_dtEndDate = dtEndDate
    End Sub

    Private Sub fPrintDialog_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        mp_lXMargin = 50
        mp_lYMargin = 50
        m_lXPreviewMargin = 100
        m_lYPreviewMargin = 100

        If mp_dtStartDate.DateTimePart.Ticks = 0 Then
            mp_dtStartDate = mp_oControl.CurrentViewObject.TimeLine.StartDate
        End If
        If mp_dtEndDate.DateTimePart.Ticks = 0 Then
            mp_dtEndDate = mp_oControl.CurrentViewObject.TimeLine.EndDate
        End If

        txtSDYear.Text = mp_dtStartDate.Year.ToString()
        txtSDMonth.Text = mp_dtStartDate.Month.ToString()
        txtSDDay.Text = mp_dtStartDate.Day.ToString()
        txtSDHour.Text = mp_dtStartDate.Hour.ToString()
        txtSDMinute.Text = mp_dtStartDate.Minute.ToString()
        txtSDSecond.Text = mp_dtStartDate.Second.ToString()

        txtEDYear.Text = mp_dtEndDate.Year.ToString()
        txtEDMonth.Text = mp_dtEndDate.Month.ToString()
        txtEDDay.Text = mp_dtEndDate.Day.ToString()
        txtEDHour.Text = mp_dtEndDate.Hour.ToString()
        txtEDMinute.Text = mp_dtEndDate.Minute.ToString()
        txtEDSecond.Text = mp_dtEndDate.Second.ToString()

        txtPageHeight.Text = "920"
        txtPageWidth.Text = "692"
        txtScale.Text = 100
        txtStartPage.Text = 1
        txtEndPage.Text = TotalPages

        mp_bLoaded = True
        Recalculate()


    End Sub

    Public ReadOnly Property StartPage() As Integer
        Get
            If txtStartPage.Text.Length = 0 Then
                Return 1
                Exit Property
            End If
            Return txtStartPage.Text
        End Get
    End Property

    Public ReadOnly Property EndPage() As Integer
        Get
            If txtEndPage.Text.Length = 0 Then
                Return 1
                Exit Property
            End If
            Return txtEndPage.Text
        End Get
    End Property

    Public ReadOnly Property PageWidth() As Integer
        Get
            If txtPageWidth.Text.Length = 0 Then
                Return 0
                Exit Property
            End If
            Return txtPageWidth.Text
        End Get
    End Property

    Public ReadOnly Property PageHeight() As Integer
        Get
            If txtPageHeight.Text.Length = 0 Then
                Return 0
                Exit Property
            End If
            Return txtPageHeight.Text
        End Get
    End Property

    Public ReadOnly Property PagesInXDirection() As Integer
        Get
            If PageWidth = 0 Then
                Return 0
                Exit Property
            End If
            Return System.Math.Abs(Int(-(mp_oControl.Printer.PrintAreaWidth / (PageWidth * (100 / Scale)))))
        End Get
    End Property

    Public ReadOnly Property PagesInYDirection() As Integer
        Get
            If PageHeight = 0 Then
                Return 0
                Exit Property
            End If
            Return System.Math.Abs(Int(-(mp_oControl.Printer.PrintAreaHeight / (PageHeight * (100 / Scale)))))
        End Get
    End Property

    Public ReadOnly Property TotalPages() As Integer
        Get
            Return PagesInXDirection * PagesInYDirection
        End Get
    End Property

    Public ReadOnly Property Scale() As Integer
        Get
            If txtScale.Text.Length = 0 Then
                Return 100
                Exit Property
            End If
            Return txtScale.Text
        End Get
    End Property

    Public Property XMargin() As Integer
        Get
            Return mp_lXMargin
        End Get
        Set(ByVal vNewValue As Integer)
            mp_lXMargin = vNewValue
        End Set
    End Property

    Public Property YMargin() As Integer
        Get
            Return mp_lYMargin
        End Get
        Set(ByVal vNewValue As Integer)
            mp_lYMargin = vNewValue
        End Set
    End Property

    Public Sub PrintControl(ByRef oVisual As Visual, ByVal lPageNumber As Integer, Optional ByVal lPhysicalOffsetX As Long = 0, Optional ByVal lPhysicalOffsetY As Long = 0, Optional ByVal lHScrollPos As Long = 0, Optional ByVal lVScrollPos As Long = 0)
        mp_lRow = System.Math.Abs(Int(-(lPageNumber / PagesInXDirection)))
        mp_lColumn = lPageNumber - ((mp_lRow - 1) * PagesInXDirection)
        If ((mp_lColumn - 1) * PageWidth) + PageWidth > mp_oControl.Printer.PrintAreaWidth Then
            mp_lPageWidth = mp_oControl.Printer.PrintAreaWidth - ((mp_lColumn - 1) * PageWidth)
        Else
            mp_lPageWidth = PageWidth
        End If
        If ((mp_lRow - 1) * PageHeight) + PageHeight > mp_oControl.Printer.PrintAreaHeight Then
            mp_lPageHeight = mp_oControl.Printer.PrintAreaHeight - ((mp_lRow - 1) * PageHeight)
        Else
            mp_lPageHeight = PageHeight
        End If
        mp_oControl.Printer.PrintControl(oVisual, (mp_lColumn - 1) * PageWidth, (mp_lRow - 1) * PageHeight, mp_lPageWidth, mp_lPageHeight, mp_lXMargin, mp_lYMargin, Scale)
    End Sub

    Private Sub txtPageWidth_TextChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.TextChangedEventArgs) Handles txtPageWidth.TextChanged
        Recalculate()
    End Sub

    Private Sub txtPageHeight_TextChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.TextChangedEventArgs) Handles txtPageHeight.TextChanged
        Recalculate()
    End Sub

    Private Sub txtScale_TextChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.TextChangedEventArgs) Handles txtScale.TextChanged
        Recalculate()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdPrint.Click
        If InitializePrinter() = True Then
            Dim oForm As New PrintDialog
            oForm.PageRangeSelection = PageRangeSelection.AllPages
            oForm.UserPageRangeEnabled = True
            If oForm.ShowDialog() = True Then
                Dim oDocSequence As FixedDocumentSequence = GetXPSDocument.GetFixedDocumentSequence()
                If Not oDocSequence Is Nothing Then
                    oForm.PrintDocument(oDocSequence.DocumentPaginator, "ActiveGantt")
                End If
            End If
        End If
    End Sub

    Private Sub cmdSaveXPS_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdSaveXPS.Click
        If InitializePrinter() = True Then
            Dim oForm As New Microsoft.Win32.SaveFileDialog
            oForm.FileName = "ActiveGantt"
            oForm.DefaultExt = ".xps"
            oForm.Filter = "XPS documents (.xps)|*.xps"
            If oForm.ShowDialog() = True Then
                SaveXPSToFile(oForm.FileName)
            End If
        End If
    End Sub

    Private Sub cmdPreview_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdPreview.Click
        If InitializePrinter() = True Then
            Dim oForm As New fPrintPreview
            oForm.oXPSDoc = GetXPSDocument()
            If Not oForm.oXPSDoc Is Nothing Then
                oForm.oURI = oURI
                oForm.ShowDialog()
            End If
        End If
    End Sub

    Private Sub SaveXPSToFile(ByVal sFilename As String)
        Dim oPackage As Package = Package.Open(sFilename, FileMode.Create)
        Dim oXPSDoc As New XpsDocument(oPackage)
        Dim oWriter As Xps.XpsDocumentWriter = XpsDocument.CreateXpsDocumentWriter(oXPSDoc)
        Dim i As Integer
        Dim oMultipleVisualsDoc As Xps.VisualsToXpsDocument = oWriter.CreateVisualsCollator()
        For i = StartPage To EndPage
            Dim oVisual As Visual = Nothing
            PrintControl(oVisual, i)
            oMultipleVisualsDoc.Write(oVisual)
        Next
        oMultipleVisualsDoc.EndBatchWrite()
        oXPSDoc.Close()
        oPackage.Close()
    End Sub

    Public Function GetXPSDocument() As XpsDocument
        If EndPage >= StartPage Then
            oURI = New Uri(sURI)
            Dim oStream As New System.IO.MemoryStream()
            Dim oPackage As Package = Package.Open(oStream, FileMode.Create)
            PackageStore.AddPackage(oURI, oPackage)
            Dim oXPSDoc As New XpsDocument(oPackage, CompressionOption.Maximum, sURI)
            Dim oWriter As Xps.XpsDocumentWriter = XpsDocument.CreateXpsDocumentWriter(oXPSDoc)
            Dim i As Integer
            Dim oMultipleVisualsDoc As Xps.VisualsToXpsDocument = oWriter.CreateVisualsCollator()
            For i = StartPage To EndPage
                Dim oVisual As Visual = Nothing
                PrintControl(oVisual, i)
                oMultipleVisualsDoc.Write(oVisual)
            Next
            oMultipleVisualsDoc.EndBatchWrite()
            Return oXPSDoc
        Else
            Return Nothing
        End If
    End Function

    Private Function InitializePrinter() As Boolean
        mp_dtStartDate = New AGVBW.DateTime(GetTextNum(txtSDYear, 0), GetTextNum(txtSDMonth, 0), GetTextNum(txtSDDay, 0), GetTextNum(txtSDHour, 0), GetTextNum(txtSDMinute, 0), GetTextNum(txtSDSecond, 0))
        mp_dtEndDate = New AGVBW.DateTime(GetTextNum(txtEDYear, 0), GetTextNum(txtEDMonth, 0), GetTextNum(txtEDDay, 0), GetTextNum(txtEDHour, 0), GetTextNum(txtEDMinute, 0), GetTextNum(txtEDSecond, 0))
        If mp_dtStartDate.DateTimePart.Ticks = 0 Then
            Return False
        End If
        If mp_dtEndDate.DateTimePart.Ticks = 0 Then
            Return False
        End If
        If mp_dtEndDate <= mp_dtStartDate Then
            MessageBox.Show("The end date cannot be smaller than or equal to the start date.")
            Return False
        End If
        If PageWidth = 0 Then
            MessageBox.Show("The page width must be greater than zero.")
            Return False
        End If
        If PageHeight = 0 Then
            MessageBox.Show("The page height must be greater than zero.")
            Return False
        End If
        mp_oControl.Printer.Initialize(mp_dtStartDate, mp_dtEndDate, -1)
        Return True
    End Function

    Private Sub Recalculate()
        If (mp_bLoaded = False) Then
            Return
        End If
        If (InitializePrinter() = True) Then
            lblNumberOfPages.Content = "Total Pages: " & TotalPages
            txtEndPage.Text = TotalPages
        End If
    End Sub


    Private Sub txtSDYear_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs) Handles txtSDYear.TextChanged
        Recalculate()
    End Sub

    Private Sub txtSDMonth_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs) Handles txtSDMonth.TextChanged
        Recalculate()
    End Sub

    Private Sub txtSDDay_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs)
        Recalculate()
    End Sub

    Private Sub txtSDDay_TextChanged_1(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs) Handles txtSDDay.TextChanged
        Recalculate()
    End Sub

    Private Sub txtSDHour_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs) Handles txtSDHour.TextChanged
        Recalculate()
    End Sub

    Private Sub txtSDMinute_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs)
        Recalculate()
    End Sub

    Private Sub txtSDSecond_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs) Handles txtSDSecond.TextChanged
        Recalculate()
    End Sub

    Private Sub txtEDYear_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs) Handles txtEDYear.TextChanged
        Recalculate()
    End Sub

    Private Sub txtEDMonth_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs) Handles txtEDMonth.TextChanged
        Recalculate()
    End Sub

    Private Sub txtEDDay_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs) Handles txtEDDay.TextChanged
        Recalculate()
    End Sub

    Private Sub txtEDHour_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs) Handles txtEDHour.TextChanged
        Recalculate()
    End Sub

    Private Sub txtEDMinute_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs) Handles txtEDMinute.TextChanged
        Recalculate()
    End Sub

    Private Sub txtEDSecond_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs) Handles txtEDSecond.TextChanged
        Recalculate()
    End Sub
End Class
