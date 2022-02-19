Imports AGVBW

Partial Public Class fLoadXML

    Private bLoaded As Boolean = False

    Private Sub Window1_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Window1.Loaded
        Me.WindowState = Windows.WindowState.Maximized
    End Sub

    Private Sub mnuLoadXML_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuLoadXML.Click
        LoadXML()
    End Sub

    Private Sub mnuSaveXML_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuSaveXML.Click
        SaveXML()
    End Sub

    Private Sub cmdLoadXML_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdLoadXML.Click
        LoadXML()
    End Sub

    Private Sub cmdSaveXML_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdSaveXML.Click
        SaveXML()
    End Sub

    Private Sub mnuClose_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuClose.Click
        Me.Close()
    End Sub

    Private Sub LoadXML()
        Dim dlg As New Microsoft.Win32.OpenFileDialog()
        dlg.DefaultExt = ".xml"
        dlg.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
        If dlg.ShowDialog() = True Then
            ActiveGanttVBWCtl1.ReadXML(dlg.FileName)
            bLoaded = True
            ActiveGanttVBWCtl1.Redraw()
        End If
    End Sub

    Private Sub SaveXML()
        Dim dlg As New Microsoft.Win32.SaveFileDialog()
        If ActiveGanttVBWCtl1.ControlTag = "WBSProject" Then
            dlg.FileName = "AGVBW_WBSP"
        ElseIf ActiveGanttVBWCtl1.ControlTag = "CarRental" Then
            dlg.FileName = "AGVBW_CR"
        End If
        dlg.DefaultExt = ".xml"
        dlg.Filter = "XML Files (.xml)|*.xml"
        If dlg.ShowDialog() = True Then
            ActiveGanttVBWCtl1.WriteXML(dlg.FileName)
        End If
    End Sub

    Private Sub ActiveGanttVBWCtl1_CustomTierDraw(ByVal sender As Object, ByVal e As AGVBW.CustomTierDrawEventArgs) Handles ActiveGanttVBWCtl1.CustomTierDraw
        If bLoaded = False Then
            Return
        End If
        If ActiveGanttVBWCtl1.ControlTag = "WBSProject" Then
            If e.TierPosition = E_TIERPOSITION.SP_LOWER Then
                e.StyleIndex = "TimeLineTiers"
                e.Text = e.StartDate.ToString("MMM")
            ElseIf e.TierPosition = E_TIERPOSITION.SP_UPPER Then
                e.StyleIndex = "TimeLineTiers"
                e.Text = e.StartDate.Year & " Q" & e.StartDate.Quarter
            End If
        ElseIf ActiveGanttVBWCtl1.ControlTag = "CarRental" Then
            If e.Interval = E_INTERVAL.IL_HOUR And e.Factor = 12 Then
                e.Text = e.StartDate.ToString("tt").ToUpper()
                e.StyleIndex = "TimeLine"
            End If
            If e.Interval = E_INTERVAL.IL_MONTH And e.Factor = 1 Then
                e.Text = e.StartDate.ToString("MMMM yyyy")
                e.StyleIndex = "TimeLineVA"
            End If
            If e.Interval = E_INTERVAL.IL_DAY And e.Factor = 1 Then
                e.Text = e.StartDate.ToString("ddd d")
                e.StyleIndex = "TimeLine"
            End If
        End If
    End Sub

    Private Sub Window1_SizeChanged(ByVal sender As Object, ByVal e As System.Windows.SizeChangedEventArgs) Handles Window1.SizeChanged
        ResizeAG()
    End Sub

    Private Sub Window1_StateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Window1.StateChanged
        ResizeAG()
    End Sub

    Private Sub ResizeAG()
        If Me.WindowState = Windows.WindowState.Normal Or Me.WindowState = Windows.WindowState.Maximized Then
            ActiveGanttVBWCtl1.Width = AGContainerGrid.ActualWidth
            ActiveGanttVBWCtl1.Height = AGContainerGrid.ActualHeight
        End If
    End Sub



End Class
