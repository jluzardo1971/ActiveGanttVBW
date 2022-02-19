Imports AGVBW

Public Class fFastLoading

    Private Sub ActiveGanttVBWCtl1_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles ActiveGanttVBWCtl1.Loaded
        Dim i As Integer
        ActiveGanttVBWCtl1.Columns.Add("Tasks", "", 200, "")
        ActiveGanttVBWCtl1.TreeviewColumnIndex = 1
        ActiveGanttVBWCtl1.Rows.BeginLoad(False)
        ActiveGanttVBWCtl1.Tasks.BeginLoad(False)
        Dim lCurrentDepth As Integer = 0
        For i = 0 To 5000
            Dim oRow As clsRow
            Dim oTask As clsTask
            oRow = ActiveGanttVBWCtl1.Rows.Load("K" & i.ToString)
            oTask = ActiveGanttVBWCtl1.Tasks.Load("K" & i.ToString(), "K" & i.ToString)
            oRow.Text = "Task K" & i.ToString()
            oRow.MergeCells = True
            oRow.Node.Depth = lCurrentDepth
            oTask.Text = "K" & i.ToString()
            oTask.StartDate = ActiveGanttVBWCtl1.MathLib.DateTimeAdd(E_INTERVAL.IL_HOUR, GetRnd(0, 5), AGVBW.DateTime.Now)
            oTask.EndDate = ActiveGanttVBWCtl1.MathLib.DateTimeAdd(E_INTERVAL.IL_HOUR, GetRnd(2, 7), oTask.StartDate)
            lCurrentDepth = lCurrentDepth + GetRnd(-1, 2)
            If lCurrentDepth < 0 Then
                lCurrentDepth = 0
            End If
        Next
        ActiveGanttVBWCtl1.Tasks.EndLoad()
        ActiveGanttVBWCtl1.Rows.EndLoad()
        ActiveGanttVBWCtl1.Rows.BeginLoad(True)
        ActiveGanttVBWCtl1.Tasks.BeginLoad(True)
        For i = 5001 To 10000
            Dim oRow As clsRow
            Dim oTask As clsTask
            oRow = ActiveGanttVBWCtl1.Rows.Load("KL" & i.ToString)
            oTask = ActiveGanttVBWCtl1.Tasks.Load("KL" & i.ToString, "KL" & i.ToString)
            oRow.Text = "Task KL" & i.ToString()
            oRow.MergeCells = True
            oTask.Text = "KL" & i.ToString()
            oTask.StartDate = ActiveGanttVBWCtl1.MathLib.DateTimeAdd(E_INTERVAL.IL_HOUR, GetRnd(0, 5), AGVBW.DateTime.Now)
            oTask.EndDate = ActiveGanttVBWCtl1.MathLib.DateTimeAdd(E_INTERVAL.IL_HOUR, GetRnd(2, 7), oTask.StartDate)
        Next
        ActiveGanttVBWCtl1.Tasks.EndLoad()
        ActiveGanttVBWCtl1.Rows.EndLoad()
        ActiveGanttVBWCtl1.Redraw()
    End Sub
End Class
