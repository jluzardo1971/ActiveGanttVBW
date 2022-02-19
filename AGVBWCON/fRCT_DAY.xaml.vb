Imports AGVBW

Public Class fRCT_DAY

    Private Sub Window_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        ActiveGanttVBWCtl1.Visibility = Windows.Visibility.Hidden
        Me.WindowState = Windows.WindowState.Maximized

        Dim oView As clsView
        oView = ActiveGanttVBWCtl1.Views.Add(E_INTERVAL.IL_MINUTE, 10, E_TIERTYPE.ST_MONTH, E_TIERTYPE.ST_DAYOFWEEK, E_TIERTYPE.ST_DAYOFWEEK, "View1")
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TickMarkArea.Visible = False

        ActiveGanttVBWCtl1.CurrentView = "View1"

        Dim i As Integer
        For i = 1 To 50
            ActiveGanttVBWCtl1.Rows.Add("K" & i.ToString())
        Next

        Dim oTimeBlock As clsTimeBlock

        oTimeBlock = ActiveGanttVBWCtl1.TimeBlocks.Add("TB_OutOfOfficeHours")
        oTimeBlock.NonWorking = True
        oTimeBlock.BaseDate = New AGVBW.DateTime(2000, 1, 1, 18, 0, 0)
        oTimeBlock.DurationInterval = E_INTERVAL.IL_HOUR
        oTimeBlock.DurationFactor = 13
        oTimeBlock.TimeBlockType = E_TIMEBLOCKTYPE.TBT_RECURRING
        oTimeBlock.RecurringType = E_RECURRINGTYPE.RCT_DAY

        oTimeBlock = ActiveGanttVBWCtl1.TimeBlocks.Add("TB_LunchBreak")
        oTimeBlock.NonWorking = True
        oTimeBlock.BaseDate = New AGVBW.DateTime(2000, 1, 1, 12, 0, 0)
        oTimeBlock.DurationInterval = E_INTERVAL.IL_MINUTE
        oTimeBlock.DurationFactor = 90
        oTimeBlock.TimeBlockType = E_TIMEBLOCKTYPE.TBT_RECURRING
        oTimeBlock.RecurringType = E_RECURRINGTYPE.RCT_DAY


        ActiveGanttVBWCtl1.Width = AGContainerGrid.ActualWidth
        ActiveGanttVBWCtl1.Height = AGContainerGrid.ActualHeight

        ActiveGanttVBWCtl1.Visibility = Windows.Visibility.Visible
    End Sub
End Class
