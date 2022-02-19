Imports AGVBW

Public Class fSortRows

    Private mp_bDescending As Boolean = False

    Private Sub Window_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        Dim i As Integer
        ActiveGanttVBWCtl1.Columns.Add("", "C1", 125, "")
        For i = 1 To 10
            Dim si As String
            si = i.ToString
            While si.Length < 2
                si = "0" & si
            End While
            ActiveGanttVBWCtl1.Rows.Add("K" & si, "K" & si, True, True, "")
        Next
    End Sub

    Private Sub cmdSortRows_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdSortRows.Click
        mp_bDescending = Not mp_bDescending
        ActiveGanttVBWCtl1.Rows.SortRows("Text", mp_bDescending, E_SORTTYPE.ES_STRING, -1, -1)
        ActiveGanttVBWCtl1.Redraw()
    End Sub
End Class
