Public Class fWBSPProperties

    Dim mp_oParent As fWBSProject

    Public Sub New(ByVal oParent As fWBSProject)
        InitializeComponent()
        mp_oParent = oParent
    End Sub

    Private Sub fWBSPProperties_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        chkEnforcePredecessors.IsChecked = mp_oParent.ActiveGanttVBWCtl1.EnforcePredecessors
        cboPredecessorMode.SelectedValue = System.Convert.ToInt32(mp_oParent.ActiveGanttVBWCtl1.PredecessorMode).ToString()
    End Sub

    Private Sub cmdOK_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdOK.Click
        mp_oParent.ActiveGanttVBWCtl1.EnforcePredecessors = chkEnforcePredecessors.IsChecked
        mp_oParent.ActiveGanttVBWCtl1.PredecessorMode = DirectCast(System.Convert.ToInt32(cboPredecessorMode.SelectedValue), AGVBW.E_PREDECESSORMODE)
        mp_oParent.ActiveGanttVBWCtl1.Redraw()
        Me.Close()
    End Sub

End Class
