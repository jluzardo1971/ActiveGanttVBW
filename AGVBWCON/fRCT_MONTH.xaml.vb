Imports AGVBW

Public Class fRCT_MONTH

    Private Sub Window_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        ActiveGanttVBWCtl1.Visibility = Windows.Visibility.Hidden
        Me.WindowState = Windows.WindowState.Maximized

        Dim oView As clsView
        oView = ActiveGanttVBWCtl1.Views.Add(E_INTERVAL.IL_HOUR, 1, E_TIERTYPE.ST_MONTH, E_TIERTYPE.ST_DAYOFWEEK, E_TIERTYPE.ST_DAYOFWEEK, "View1")
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TickMarkArea.Visible = False
        ActiveGanttVBWCtl1.TierFormatScope = E_TIERFORMATSCOPE.TFS_CONTROL
        ActiveGanttVBWCtl1.TierFormat.DayOfWeekIntervalFormat = "dd"

        ActiveGanttVBWCtl1.CurrentView = "View1"

        Dim i As Integer
        For i = 1 To 50
            ActiveGanttVBWCtl1.Rows.Add("K" & i.ToString())
        Next

        Dim oTimeBlock As clsTimeBlock
        Dim dtDate As AGVBW.DateTime
        dtDate = New AGVBW.DateTime(2000, 1, 1, 0, 0, 0)

        oTimeBlock = ActiveGanttVBWCtl1.TimeBlocks.Add("TimeBlock1")
        oTimeBlock.BaseDate = dtDate
        oTimeBlock.DurationInterval = E_INTERVAL.IL_HOUR
        oTimeBlock.DurationFactor = -48
        oTimeBlock.TimeBlockType = E_TIMEBLOCKTYPE.TBT_RECURRING
        oTimeBlock.RecurringType = E_RECURRINGTYPE.RCT_MONTH


        ActiveGanttVBWCtl1.Width = AGContainerGrid.ActualWidth
        ActiveGanttVBWCtl1.Height = AGContainerGrid.ActualHeight

        ActiveGanttVBWCtl1.Visibility = Windows.Visibility.Visible
    End Sub
End Class
