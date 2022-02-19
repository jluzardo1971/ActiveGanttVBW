Imports AGVBW

Public Class fMillisecondInterval

    Private Sub Window_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        Dim oView As clsView
        oView = ActiveGanttVBWCtl1.Views.Add(E_INTERVAL.IL_MILLISECOND, 5, E_TIERTYPE.ST_MINUTE, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_SECOND, "MSI")
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.TierFormat.MinuteIntervalFormat = "MMM dd, yyyy hh:mm tt"
        oView.TimeLine.Position(New AGVBW.DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, System.DateTime.Now.Day, System.DateTime.Now.Hour, System.DateTime.Now.Minute, 58))

        ActiveGanttVBWCtl1.CurrentView = "MSI"

        Dim i As Integer
        ActiveGanttVBWCtl1.Columns.Add("", "C1", 125, "")
        For i = 1 To 10
            Dim oRow As clsRow
            oRow = ActiveGanttVBWCtl1.Rows.Add("K" & i.ToString, "K" & i.ToString(), True, True, "")
        Next
    End Sub

    Private Sub ActiveGanttVBWCtl1_CompleteObjectMove(sender As System.Object, e As AGVBW.ObjectStateChangedEventArgs) Handles ActiveGanttVBWCtl1.CompleteObjectMove
        If e.EventTarget = E_EVENTTARGET.EVT_TASK Then
            Dim oTask As clsTask
            Dim sText As String
            oTask = ActiveGanttVBWCtl1.Tasks.Item(e.Index.ToString())
            sText = ActiveGanttVBWCtl1.MathLib.DateTimeDiff(E_INTERVAL.IL_MILLISECOND, oTask.StartDate, oTask.EndDate).ToString()
            oTask.Text = sText & "ms"
        End If
    End Sub

    Private Sub ActiveGanttVBWCtl1_CompleteObjectSize(sender As System.Object, e As AGVBW.ObjectStateChangedEventArgs) Handles ActiveGanttVBWCtl1.CompleteObjectSize
        If e.EventTarget = E_EVENTTARGET.EVT_TASK Then
            Dim oTask As clsTask
            Dim sText As String
            oTask = ActiveGanttVBWCtl1.Tasks.Item(e.Index.ToString())
            sText = ActiveGanttVBWCtl1.MathLib.DateTimeDiff(E_INTERVAL.IL_MILLISECOND, oTask.StartDate, oTask.EndDate).ToString()
            oTask.Text = sText & "ms"
        End If
    End Sub

    Private Sub ActiveGanttVBWCtl1_ObjectAdded(sender As System.Object, e As AGVBW.ObjectAddedEventArgs) Handles ActiveGanttVBWCtl1.ObjectAdded
        If e.EventTarget = E_EVENTTARGET.EVT_TASK Then
            Dim oTask As clsTask
            Dim sText As String
            oTask = ActiveGanttVBWCtl1.Tasks.Item(e.TaskIndex.ToString())
            sText = ActiveGanttVBWCtl1.MathLib.DateTimeDiff(E_INTERVAL.IL_MILLISECOND, oTask.StartDate, oTask.EndDate).ToString()
            oTask.Text = sText & "ms"
        End If
    End Sub
End Class
