Imports System.Data.OleDb
Imports AGVBW
Imports System.Data

Partial Public Class fWBSProjectTaskView
    Private mp_lTaskIndex As Integer
    Private mp_oParent As fWBSProject
    Private mp_oGrid As Grid

    Friend Sub New(ByRef oParent As fWBSProject, ByVal lTaskIndex As Integer)
        InitializeComponent()
        mp_oParent = oParent
        mp_lTaskIndex = lTaskIndex
    End Sub

    Private Sub fWBSProjectTaskView(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        Dim oRows() As DataRow
        Dim oColumn1 As New ColumnDefinition
        oColumn1.Width = New System.Windows.GridLength(40)
        Dim oColumn2 As New ColumnDefinition
        oColumn2.Width = New System.Windows.GridLength(260)
        Dim oColumn3 As New ColumnDefinition
        oColumn3.Width = New System.Windows.GridLength(100)
        Dim oRow As New RowDefinition
        oRow.Height = New System.Windows.GridLength(20)
        Dim i As Integer = 1
        Dim j As Integer = 0

        Dim oTask As clsTask
        oTask = mp_oParent.ActiveGanttVBWCtl1.Tasks.Item(mp_lTaskIndex)
        Me.Title = oTask.Row.Text
        txtTaskName.Text = oTask.Row.Text

        mp_oGrid = New Grid()
        mp_oGrid.ColumnDefinitions.Add(oColumn1)
        mp_oGrid.ColumnDefinitions.Add(oColumn2)
        mp_oGrid.ColumnDefinitions.Add(oColumn3)
        mp_oGrid.RowDefinitions.Add(oRow)
        AddTextBlock("ID", 0, 0)
        AddTextBlock("Predecessor Task Name", 0, 1)
        AddTextBlock("Type", 0, 2)
        If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
                Dim oReader As OleDbDataReader
                oReader = DST_ACCESS.g_DST_ACCESS_ReturnReader("SELECT * FROM qry_GuysStThomas_Predecessors WHERE lSuccessorID=" & mp_lTaskIndex, oConn)
                While oReader.Read = True
                    Dim oRowD As New RowDefinition
                    oRowD.Height = New System.Windows.GridLength(20)
                    mp_oGrid.RowDefinitions.Add(oRowD)
                    AddTextBlock(oReader.Item("lPredecessorID"), i, 0)
                    AddTextBlock(oReader.Item("sDescription"), i, 1)
                    AddTextBlock(oReader.Item("sPredecessorType"), i, 2)
                    i = i + 1
                End While
                oReader.Close()
            End Using
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            oRows = mp_oParent.mp_otb_GuysStThomas_Predecessors.Tables(1).Select("lSuccessorID = " & mp_lTaskIndex.ToString())
            For Each oDataRow1 As DataRow In oRows
                Dim oRowD1 As New RowDefinition
                oRowD1.Height = New System.Windows.GridLength(20)
                mp_oGrid.RowDefinitions.Add(oRowD1)
                AddTextBlock(oDataRow1("lPredecessorID"), i, 0)
                AddTextBlock(mp_oParent.mp_otb_GuysStThomas.Tables(1).Rows.Find(oDataRow1("lPredecessorID")).Item("sDescription"), i, 1)
                AddTextBlock(GetPredecessorType(DirectCast(oDataRow1("yType"), System.Int32)), i, 2)
                i = i + 1
            Next
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            For j = 1 To mp_oParent.ActiveGanttVBWCtl1.Predecessors.Count
                Dim oPredecessor As clsPredecessor = mp_oParent.ActiveGanttVBWCtl1.Predecessors.Item(j.ToString())
                If oPredecessor.SuccessorKey = oTask.Key Then
                    Dim oRowD2 As New RowDefinition
                    oRowD2.Height = New System.Windows.GridLength(20)
                    mp_oGrid.RowDefinitions.Add(oRowD2)
                    AddTextBlock(oPredecessor.Key.Replace("K", ""), i, 0)
                    AddTextBlock(GetTaskDescriptionByTaskKey(oPredecessor.PredecessorKey), i, 1)
                    AddTextBlock(GetPredecessorType(oPredecessor.PredecessorType), i, 2)
                    i = i + 1
                End If
            Next
        End If
        mp_oGrid.SetValue(Canvas.LeftProperty, CDbl(8))
        mp_oGrid.SetValue(Canvas.TopProperty, CDbl(50))
        oCanvas.Children.Add(mp_oGrid)
    End Sub

    Private Sub AddTextBlock(ByVal sText As String, ByVal lRow As Integer, ByVal lColumn As Integer)
        Dim oRectangle As New Rectangle
        oRectangle.Stroke = Brushes.Black
        Select Case lColumn
            Case 0
                oRectangle.Width = 41
                oRectangle.SetValue(Canvas.LeftProperty, CDbl(6))
            Case 1
                oRectangle.Width = 261
                oRectangle.SetValue(Canvas.LeftProperty, CDbl(46))
            Case 2
                oRectangle.Width = 101
                oRectangle.SetValue(Canvas.LeftProperty, CDbl(306))
        End Select
        oRectangle.SetValue(Canvas.TopProperty, CDbl(50 + (19 * lRow)))
        oRectangle.Height = 20
        Dim oTextBlock As New TextBlock
        oTextBlock.Text = sText
        oTextBlock.FontSize = 11
        If lRow = 0 Then
            oRectangle.Fill = Brushes.Gray
            oTextBlock.FontWeight = FontWeights.Bold
        Else
            oRectangle.Fill = Brushes.White
        End If
        Grid.SetRow(oTextBlock, lRow)
        Grid.SetColumn(oTextBlock, lColumn)
        oCanvas.Children.Add(oRectangle)
        mp_oGrid.Children.Add(oTextBlock)
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdOK.Click
        Me.Close()
    End Sub

    Private Function GetPredecessorType(ByVal lType As Integer) As String
        If lType = 0 Then
            Return "End-To-Start (ES)"
        ElseIf lType = 1 Then
            Return "Start-To-Start (SS)"
        ElseIf lType = 2 Then
            Return "End-To-End (EE)"
        ElseIf lType = 3 Then
            Return "Start-To-End (SE)"
        End If
        Return ""
    End Function

    Private Function GetTaskDescriptionByTaskKey(ByVal sTaskKey As String) As String
        Dim i As Integer = 0
        For i = 1 To mp_oParent.ActiveGanttVBWCtl1.Tasks.Count
            If mp_oParent.ActiveGanttVBWCtl1.Tasks.Item(i.ToString()).Key = sTaskKey Then
                Return mp_oParent.ActiveGanttVBWCtl1.Rows.Item(mp_oParent.ActiveGanttVBWCtl1.Tasks.Item(i.ToString()).RowKey).Node.Text
            End If
        Next
        Return ""
    End Function
End Class
