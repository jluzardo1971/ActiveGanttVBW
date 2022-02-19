Imports AGVBW

Public Class fCustomDrawing

    Private Sub ActiveGanttVBWCtl1_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles ActiveGanttVBWCtl1.Loaded
        Dim i As Integer
        ActiveGanttVBWCtl1.Columns.Add("Column 1", "", 125, "")
        ActiveGanttVBWCtl1.Columns.Add("Column 2", "", 100, "")
        For i = 1 To 10
            ActiveGanttVBWCtl1.Rows.Add("K" & i.ToString(), "Row " & i.ToString() & " (Key: " & "K" & i.ToString() & ")", True, True, "")
        Next

        ActiveGanttVBWCtl1.CurrentViewObject.TimeLine.Position(New AGVBW.DateTime(2011, 11, 21, 0, 0, 0))
        ActiveGanttVBWCtl1.Tasks.Add("Task 1", "K1", New AGVBW.DateTime(2011, 11, 21, 0, 0, 0), New AGVBW.DateTime(2011, 11, 21, 3, 0, 0), "", "", "")
        ActiveGanttVBWCtl1.Tasks.Add("Task 2", "K2", New AGVBW.DateTime(2011, 11, 21, 1, 0, 0), New AGVBW.DateTime(2011, 11, 21, 4, 0, 0), "", "", "")
        ActiveGanttVBWCtl1.Tasks.Add("Task 3", "K3", New AGVBW.DateTime(2011, 11, 21, 2, 0, 0), New AGVBW.DateTime(2011, 11, 21, 5, 0, 0), "", "", "")

        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub ActiveGanttVBWCtl1_Draw(sender As System.Object, e As AGVBW.DrawEventArgs) Handles ActiveGanttVBWCtl1.Draw
        If e.EventTarget = E_EVENTTARGET.EVT_TASK Then
            If ActiveGanttVBWCtl1.SelectedTaskIndex = e.ObjectIndex Then
                e.CustomDraw = True
                Dim oTask As clsTask
                oTask = ActiveGanttVBWCtl1.Tasks.Item(e.ObjectIndex.ToString())
                Dim oFont As New Font("Arial", 7, FontWeights.Normal)
                Dim oTextFlags As New clsTextFlags(ActiveGanttVBWCtl1)
                oTextFlags.HorizontalAlignment = GRE_HORIZONTALALIGNMENT.HAL_CENTER
                oTextFlags.VerticalAlignment = GRE_VERTICALALIGNMENT.VAL_CENTER
                Dim oImage As New Image()
                Dim oURI As New Uri("../Images/sampleimage.jpg", UriKind.RelativeOrAbsolute)
                Dim oDecoder As New JpegBitmapDecoder(oURI, BitmapCreateOptions.None, BitmapCacheOption.None)
                Dim oBitmap As BitmapSource = oDecoder.Frames(0)
                oImage.Source = oBitmap
                ActiveGanttVBWCtl1.Drawing.PaintImage(oImage, oTask.Left + 40, oTask.Top + 10, oTask.Left + 64, oTask.Top + 34)
                ActiveGanttVBWCtl1.Drawing.DrawLine(oTask.Left, ((oTask.Bottom - oTask.Top) / 2) + oTask.Top, oTask.Right, ((oTask.Bottom - oTask.Top) / 2) + oTask.Top, Colors.Green, GRE_LINEDRAWSTYLE.LDS_SOLID, 1)
                ActiveGanttVBWCtl1.Drawing.DrawRectangle(oTask.Left, oTask.Top, oTask.Left + 10, oTask.Top + 10, Colors.Red, GRE_LINEDRAWSTYLE.LDS_SOLID, 1)
                ActiveGanttVBWCtl1.Drawing.DrawBorder(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom, Colors.Blue, GRE_LINEDRAWSTYLE.LDS_SOLID, 2)
                ActiveGanttVBWCtl1.Drawing.DrawAlignedText(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom, oTask.Text & " Is Selected", GRE_HORIZONTALALIGNMENT.HAL_RIGHT, GRE_VERTICALALIGNMENT.VAL_BOTTOM, Colors.Blue, oFont)
                ActiveGanttVBWCtl1.Drawing.DrawText(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom, "Draw Text", oTextFlags, Colors.Red, oFont)
            End If
        End If
    End Sub
End Class
