Option Explicit On
Imports AGVBW

Partial Public Class fMSProject11

    Private oMP11 As MSP2003.MP11
    Private mp_lControlDraw As Integer = 0
    Private mp_lControlRedrawn As Integer = 0
    Private Const mp_sFontName As String = "Tahoma"

#Region "Constructors"

#End Region

#Region "Form Loaded"

    Private Sub Window1_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        Me.Title = "The Source Code Store - ActiveGantt Scheduler Control Version " & ActiveGanttVBWCtl1.Version & " - Microsoft Project 2003 integration using XML Files and the MSP2003 Integration Library"
        Me.WindowState = Windows.WindowState.Maximized

        InitializeAG()
        ActiveGanttVBWCtl1.Redraw()
    End Sub

#End Region

#Region "Form Resizing"

    Private Sub ResizeAG()
        If Me.WindowState = Windows.WindowState.Normal Or Me.WindowState = Windows.WindowState.Maximized Then
            ActiveGanttVBWCtl1.Width = AGContainerGrid.ActualWidth
            ActiveGanttVBWCtl1.Height = AGContainerGrid.ActualHeight
        End If
    End Sub

    Private Sub fMSProject11_SizeChanged(ByVal sender As Object, ByVal e As System.Windows.SizeChangedEventArgs) Handles Me.SizeChanged
        ResizeAG()
    End Sub

    Private Sub fMSProject11_StateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.StateChanged
        ResizeAG()
    End Sub

#End Region

#Region "Functions"

    Private Sub InitializeAG()

        Dim oStyle As clsStyle = Nothing
        Dim oView As clsView = Nothing

        oStyle = ActiveGanttVBWCtl1.Styles.Add("TimeLineTiers")
        oStyle.Font = New Font(mp_sFontName, 7, System.Windows.FontWeights.Normal)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_RAISED
        oStyle.BorderColor = Colors.DarkGray
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE

        oStyle = ActiveGanttVBWCtl1.Styles.Add("TaskStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.Blue
        oStyle.BorderColor = Colors.Blue
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.TextPlacement = E_TEXTPLACEMENT.SCP_EXTERIORPLACEMENT
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT
        oStyle.TextXMargin = 10
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 10
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_HATCH
        oStyle.HatchBackColor = Colors.White
        oStyle.HatchForeColor = Colors.Blue
        oStyle.HatchStyle = GRE_HATCHSTYLE.HS_PERCENT50
        oStyle.PredecessorStyle.LineColor = Colors.Black
        oStyle.MilestoneStyle.ShapeIndex = GRE_FIGURETYPE.FT_DIAMOND
        oStyle.MilestoneStyle.FillColor = Colors.Blue
        oStyle.MilestoneStyle.BorderColor = Colors.Blue
        oStyle.PredecessorStyle.XOffset = 4
        oStyle.PredecessorStyle.YOffset = 4

        oStyle = ActiveGanttVBWCtl1.Styles.Add("SummaryStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.Green
        oStyle.BorderColor = Colors.Green
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.Visible = False
        oStyle.TextPlacement = E_TEXTPLACEMENT.SCP_EXTERIORPLACEMENT
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT
        oStyle.TextXMargin = 10
        oStyle.TaskStyle.StartShapeIndex = GRE_FIGURETYPE.FT_PROJECTDOWN
        oStyle.TaskStyle.EndShapeIndex = GRE_FIGURETYPE.FT_PROJECTDOWN
        oStyle.TaskStyle.EndFillColor = Colors.Green
        oStyle.TaskStyle.EndBorderColor = Colors.Green
        oStyle.TaskStyle.StartFillColor = Colors.Green
        oStyle.TaskStyle.StartBorderColor = Colors.Green
        oStyle.FillMode = GRE_FILLMODE.FM_UPPERHALFFILLED

        oStyle = ActiveGanttVBWCtl1.Styles.Add("CellStyleKeyColumn")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Color.FromArgb(255, 128, 128, 128)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT
        oStyle.TextXMargin = 4

        ActiveGanttVBWCtl1.AllowRowMove = True
        ActiveGanttVBWCtl1.AllowRowSize = True
        ActiveGanttVBWCtl1.AddMode = E_ADDMODE.AT_BOTH
        ActiveGanttVBWCtl1.Splitter.Position = 285
        ActiveGanttVBWCtl1.Treeview.Images = True
        ActiveGanttVBWCtl1.Treeview.CheckBoxes = True
        ActiveGanttVBWCtl1.Treeview.FullColumnSelect = True
        ActiveGanttVBWCtl1.Treeview.TreeLines = False
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.TimerInterval = 50

        Dim oColumn As clsColumn

        oColumn = ActiveGanttVBWCtl1.Columns.Add("ID", "", 30, "")
        oColumn.AllowTextEdit = True

        oColumn = ActiveGanttVBWCtl1.Columns.Add("Task Name", "", 255, "")
        oColumn.AllowTextEdit = True

        ActiveGanttVBWCtl1.TreeviewColumnIndex = 2

        oView = ActiveGanttVBWCtl1.Views.Add(E_INTERVAL.IL_HOUR, 24, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_QUARTER
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.LowerTier.Factor = 1
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TimeLineScrollBar.StartDate = AGVBW.DateTime.Now
        oView.TimeLine.TimeLineScrollBar.Enabled = True
        oView.TimeLine.TimeLineScrollBar.Visible = False

        oView = ActiveGanttVBWCtl1.Views.Add(E_INTERVAL.IL_HOUR, 12, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_QUARTER
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.LowerTier.Factor = 1
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TimeLineScrollBar.StartDate = AGVBW.DateTime.Now
        oView.TimeLine.TimeLineScrollBar.Interval = E_INTERVAL.IL_HOUR
        oView.TimeLine.TimeLineScrollBar.Factor = 1
        oView.TimeLine.TimeLineScrollBar.SmallChange = 12
        oView.TimeLine.TimeLineScrollBar.LargeChange = 240
        oView.TimeLine.TimeLineScrollBar.Max = 2000
        oView.TimeLine.TimeLineScrollBar.Value = 0
        oView.TimeLine.TimeLineScrollBar.Enabled = True
        oView.TimeLine.TimeLineScrollBar.Visible = True
        oView.ClientArea.DetectConflicts = False

        oView = ActiveGanttVBWCtl1.Views.Add(E_INTERVAL.IL_HOUR, 6, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_QUARTER
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.LowerTier.Factor = 1
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TimeLineScrollBar.StartDate = AGVBW.DateTime.Now
        oView.TimeLine.TimeLineScrollBar.Interval = E_INTERVAL.IL_HOUR
        oView.TimeLine.TimeLineScrollBar.Factor = 1
        oView.TimeLine.TimeLineScrollBar.SmallChange = 6
        oView.TimeLine.TimeLineScrollBar.LargeChange = 480
        oView.TimeLine.TimeLineScrollBar.Max = 4000
        oView.TimeLine.TimeLineScrollBar.Value = 0
        oView.TimeLine.TimeLineScrollBar.Enabled = True
        oView.TimeLine.TimeLineScrollBar.Visible = True
        oView.ClientArea.DetectConflicts = False

        oView = ActiveGanttVBWCtl1.Views.Add(E_INTERVAL.IL_HOUR, 3, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_QUARTER
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.LowerTier.Factor = 1
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TimeLineScrollBar.StartDate = AGVBW.DateTime.Now
        oView.TimeLine.TimeLineScrollBar.Interval = E_INTERVAL.IL_HOUR
        oView.TimeLine.TimeLineScrollBar.Factor = 1
        oView.TimeLine.TimeLineScrollBar.SmallChange = 3
        oView.TimeLine.TimeLineScrollBar.LargeChange = 960
        oView.TimeLine.TimeLineScrollBar.Max = 8000
        oView.TimeLine.TimeLineScrollBar.Value = 0
        oView.TimeLine.TimeLineScrollBar.Enabled = True
        oView.TimeLine.TimeLineScrollBar.Visible = True
        oView.ClientArea.DetectConflicts = False

        oView = ActiveGanttVBWCtl1.Views.Add(E_INTERVAL.IL_HOUR, 1, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_DAY
        oView.TimeLine.TierArea.LowerTier.Factor = 1
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TimeLineScrollBar.StartDate = AGVBW.DateTime.Now
        oView.TimeLine.TimeLineScrollBar.Interval = E_INTERVAL.IL_HOUR
        oView.TimeLine.TimeLineScrollBar.Factor = 1
        oView.TimeLine.TimeLineScrollBar.SmallChange = 48
        oView.TimeLine.TimeLineScrollBar.LargeChange = 2880
        oView.TimeLine.TimeLineScrollBar.Max = 24000
        oView.TimeLine.TimeLineScrollBar.Value = 0
        oView.TimeLine.TimeLineScrollBar.Enabled = True
        oView.TimeLine.TimeLineScrollBar.Visible = True
        oView.ClientArea.DetectConflicts = False

        ActiveGanttVBWCtl1.CurrentView = "5"


    End Sub

    Private Sub AGSetStartDate(ByVal dtStart As AGVBW.DateTime)
        Dim i As Integer
        For i = 1 To ActiveGanttVBWCtl1.Views.Count
            ActiveGanttVBWCtl1.Views.Item(i).TimeLine.TimeLineScrollBar.StartDate = dtStart
        Next
    End Sub

    Private Sub MP11_To_AG()
        Dim oAGTask As clsTask
        Dim oAGRow As clsRow
        Dim oMPTask As MSP2003.Task
        Dim dtStartDate As AGVBW.DateTime = AGVBW.DateTime.Now
        Dim i As Integer
        Dim j As Integer
        '// Load Project Tasks
        For i = 1 To oMP11.oTasks.Count
            oMPTask = oMP11.oTasks.Item(i)
            oAGRow = ActiveGanttVBWCtl1.Rows.Add("K" & oMPTask.lUID.ToString())
            oAGRow.Cells.Item("1").Text = oMPTask.lUID.ToString()
            oAGRow.Cells.Item("1").StyleIndex = "CellStyleKeyColumn"
            oAGRow.Height = 20
            oAGTask = ActiveGanttVBWCtl1.Tasks.Add("", "K" & oMPTask.lUID.ToString(), FromDate(oMPTask.dtStart), FromDate(oMPTask.dtFinish))
            oAGTask.Key = "K" & oMPTask.lUID.ToString()
            oAGTask.AllowedMovement = E_MOVEMENTTYPE.MT_RESTRICTEDTOROW
            oAGTask.AllowTextEdit = True
            If FromDate(oMPTask.dtStart) < dtStartDate Then
                dtStartDate = FromDate(oMPTask.dtStart)
            End If
            If oAGTask.StartDate = oAGTask.EndDate Then
                oAGTask.Text = oAGTask.StartDate.ToString("M/d")
            End If
            oAGRow.Node.Depth = oMPTask.lOutlineLevel
            oAGRow.Node.Text = oMPTask.sName
            oAGRow.Node.AllowTextEdit = True
            If oMPTask.sNotes.Length > 0 Then
                oAGRow.Node.Image = GetImage(0)
                oAGRow.Node.ImageVisible = True
            End If
        Next
        ActiveGanttVBWCtl1.Rows.UpdateTree()
        '// Indent & set Predecessors
        For i = 1 To oMP11.oTasks.Count
            oMPTask = oMP11.oTasks.Item(i)
            oAGRow = ActiveGanttVBWCtl1.Rows.Item(i)
            oAGTask = ActiveGanttVBWCtl1.Tasks.Item(i)
            If oAGRow.Node.Children > 0 Then
                oAGTask.StyleIndex = "SummaryStyle"
            Else
                oAGTask.StyleIndex = "TaskStyle"
            End If
            For j = 1 To oMPTask.oPredecessorLink_C.Count
                Dim oMPPredecessor As MSP2003.TaskPredecessorLink
                oMPPredecessor = oMPTask.oPredecessorLink_C.Item(j)
                ActiveGanttVBWCtl1.Predecessors.Add("K" & oMPTask.lUID.ToString(), "K" & oMPPredecessor.lPredecessorUID.ToString(), GetAGPredecessorType(oMPPredecessor.yType), "", "TaskStyle")
            Next
        Next
        'Assignments
        For i = 1 To oMP11.oAssignments.Count
            Dim oAssignment As MSP2003.Assignment
            oAssignment = oMP11.oAssignments.Item(i)
            oAGTask = ActiveGanttVBWCtl1.Tasks.Item("K" & oAssignment.lTaskUID)
            If oAGTask.StartDate <> oAGTask.EndDate Then
                If oAssignment.lResourceUID > 0 Then
                    If oAGTask.Text.Length = 0 Then
                        oAGTask.Text = oMP11.oResources.Item("K" & oAssignment.lResourceUID).sName
                    Else
                        oAGTask.Text = oAGTask.Text & ", " & oMP11.oResources.Item("K" & oAssignment.lResourceUID).sName
                    End If
                End If
            End If
        Next
        dtStartDate = ActiveGanttVBWCtl1.MathLib.DateTimeAdd(E_INTERVAL.IL_DAY, -3, dtStartDate)
        AGSetStartDate(dtStartDate)
    End Sub

    Private Function GetAGPredecessorType(ByVal MPPredecessorType As MSP2003.E_TYPE_3) As AGVBW.E_CONSTRAINTTYPE
        Select Case MPPredecessorType
            Case MSP2003.E_TYPE_3.T_3_FF
                Return AGVBW.E_CONSTRAINTTYPE.PCT_END_TO_END
            Case MSP2003.E_TYPE_3.T_3_FS
                Return AGVBW.E_CONSTRAINTTYPE.PCT_END_TO_START
            Case MSP2003.E_TYPE_3.T_3_SF
                Return AGVBW.E_CONSTRAINTTYPE.PCT_START_TO_END
            Case MSP2003.E_TYPE_3.T_3_SS
                Return AGVBW.E_CONSTRAINTTYPE.PCT_START_TO_START
        End Select
        Return AGVBW.E_CONSTRAINTTYPE.PCT_END_TO_START
    End Function

    Private Function GetImage(ByVal ImageIndex As Integer) As Image
        Dim oDecoder As New GifBitmapDecoder(GetURI(ImageIndex), BitmapCreateOptions.None, BitmapCacheOption.None)
        Dim oBitmap As BitmapSource = oDecoder.Frames(0)
        Dim oReturn As New Image()
        oReturn.Source = oBitmap
        Return oReturn
    End Function

    Private Function GetURI(ByVal ImageIndex As Integer) As Uri
        Dim oURI As Uri = Nothing
        Select Case ImageIndex
            Case 0
                oURI = New Uri("../images/WBS/projectnote2003.gif", UriKind.RelativeOrAbsolute)
        End Select
        Return oURI
    End Function

    Private Sub AG_To_MP11()

    End Sub

#End Region

#Region "ActiveGantt Event Handlers"

    Private Sub ActiveGanttVBWCtl1_CustomTierDraw(ByVal sender As System.Object, ByVal e As AGVBW.CustomTierDrawEventArgs) Handles ActiveGanttVBWCtl1.CustomTierDraw
        If e.TierPosition = E_TIERPOSITION.SP_UPPER Then
            e.StyleIndex = "TimeLineTiers"
            If System.Convert.ToInt32(ActiveGanttVBWCtl1.CurrentView) <= 4 Then
                e.Text = e.StartDate.Year() & " Q" & e.StartDate.Quarter()
            Else
                e.Text = e.StartDate.ToString("MMMM, yyyy")
            End If
        ElseIf e.TierPosition = E_TIERPOSITION.SP_LOWER Then
            e.StyleIndex = "TimeLineTiers"
            If System.Convert.ToInt32(ActiveGanttVBWCtl1.CurrentView) <= 4 Then
                e.Text = e.StartDate.ToString("MMM")
            Else
                e.Text = e.StartDate.ToString("ddd")
            End If
        End If
    End Sub

    Private Sub ActiveGanttVBWCtl1_ControlMouseWheel(ByVal sender As Object, ByVal e As AGVBW.MouseWheelEventArgs) Handles ActiveGanttVBWCtl1.ControlMouseWheel
        If (e.Delta = 0) Or (ActiveGanttVBWCtl1.VerticalScrollBar.Visible = False) Then
            Return
        End If
        Dim lDelta As Integer = System.Convert.ToInt32(-(e.Delta / 50))
        Dim lInitialValue As Integer = ActiveGanttVBWCtl1.VerticalScrollBar.Value
        If (ActiveGanttVBWCtl1.VerticalScrollBar.Value + lDelta < 1) Then
            ActiveGanttVBWCtl1.VerticalScrollBar.Value = 1
        ElseIf (((ActiveGanttVBWCtl1.VerticalScrollBar.Value + lDelta) > ActiveGanttVBWCtl1.VerticalScrollBar.Max)) Then
            ActiveGanttVBWCtl1.VerticalScrollBar.Value = ActiveGanttVBWCtl1.VerticalScrollBar.Max
        Else
            ActiveGanttVBWCtl1.VerticalScrollBar.Value = ActiveGanttVBWCtl1.VerticalScrollBar.Value + lDelta
        End If
        ActiveGanttVBWCtl1.Redraw()
    End Sub

#End Region

#Region "Toolbar Buttons"

    Private Sub cmdIndent_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdIndent.Click
        Indent()
    End Sub

    Private Sub cmdLoadXML_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdLoadXML.Click
        LoadXML()
    End Sub

    Private Sub cmdPrint_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdPrint.Click
        Print()
    End Sub

    Private Sub cmdSaveXML_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdSaveXML.Click
        SaveXML()
    End Sub

    Private Sub cmdZoomIn_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdZoomIn.Click
        If ActiveGanttVBWCtl1.CurrentView < ActiveGanttVBWCtl1.Views.Count Then
            ActiveGanttVBWCtl1.CurrentView = ActiveGanttVBWCtl1.CurrentView + 1
            ActiveGanttVBWCtl1.Redraw()
        End If
    End Sub

    Private Sub cmdZoomOut_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdZoomOut.Click
        If ActiveGanttVBWCtl1.CurrentView > 1 Then
            ActiveGanttVBWCtl1.CurrentView = ActiveGanttVBWCtl1.CurrentView - 1
            ActiveGanttVBWCtl1.Redraw()
        End If
    End Sub

#End Region

#Region "Menu Items"

    Private Sub mnuClose_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuClose.Click
        Me.Close()
    End Sub

    Private Sub mnuLoadXML_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuLoadXML.Click
        LoadXML()
    End Sub

    Private Sub mnuSaveXML_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuSaveXML.Click
        SaveXML()
    End Sub

    Private Sub mnuPrint_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles mnuPrint.Click
        Print()
    End Sub

#End Region

#Region "Toolbar Button & Menu Item Functions"

    Private Sub LoadXML()
        Dim OpenFileDialog1 As New Microsoft.Win32.OpenFileDialog()
        oMP11 = New MSP2003.MP11()
        OpenFileDialog1.Title = "Load MS-Project 2003 XML File"
        OpenFileDialog1.InitialDirectory = g_GetAppLocation() & "\MSP2003\"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "XML File (*.xml)|*.xml|All Files (*.*)|*.*"
        If (OpenFileDialog1.ShowDialog(Me) = True) Then
            If ValidateMSP2003(OpenFileDialog1.FileName) = False Then
                MsgBox("The file selected is not a valid Microsoft Project 2003 XML File.", MsgBoxStyle.OkOnly)
            Else
                Me.Cursor = Cursors.Wait
                ActiveGanttVBWCtl1.Clear()
                oMP11.ReadXML(OpenFileDialog1.FileName)
                Me.Cursor = Cursors.Wait
                InitializeAG()
                MP11_To_AG()
                ActiveGanttVBWCtl1.Redraw()
                ActiveGanttVBWCtl1.VerticalScrollBar.LargeChange = ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.LastVisibleRow - ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.FirstVisibleRow
                ActiveGanttVBWCtl1.Redraw()
                Me.Cursor = Cursors.Arrow
            End If
        End If
    End Sub

    Private Sub SaveXML()
        Dim SaveFileDialog1 As New Microsoft.Win32.SaveFileDialog()
        SaveFileDialog1.Title = "Save As MS-Project 2003 XML File"
        SaveFileDialog1.InitialDirectory = g_GetAppLocation() & "\MSP2003\"
        SaveFileDialog1.Filter = "XML File|*.xml"
        SaveFileDialog1.OverwritePrompt = True
        If (SaveFileDialog1.ShowDialog(Me) = True) Then
            Me.Cursor = Cursors.Wait
            AG_To_MP11()
            oMP11.WriteXML(SaveFileDialog1.FileName)
            Me.Cursor = Cursors.Arrow
        End If
    End Sub

    Private Sub Print()
        Dim oForm As New fPrintDialog(ActiveGanttVBWCtl1)
        oForm.ShowDialog()
    End Sub

    Private Sub Indent()
        Dim OpenFileDialog1 As New Microsoft.Win32.OpenFileDialog()
        Dim SaveFileDialog1 As New Microsoft.Win32.SaveFileDialog()
        OpenFileDialog1.Title = "Load XML File"
        OpenFileDialog1.InitialDirectory = g_GetAppLocation() & "\MSP2003\"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "XML File (*.xml)|*.xml|All Files (*.*)|*.*"
        If (OpenFileDialog1.ShowDialog(Me) = True) Then
            SaveFileDialog1.Title = "Save XML File As"
            SaveFileDialog1.InitialDirectory = g_GetAppLocation() & "\MSP2003\"
            SaveFileDialog1.Filter = "XML File|*.xml"
            SaveFileDialog1.OverwritePrompt = True
            If (SaveFileDialog1.ShowDialog(Me) = True) Then
                If (OpenFileDialog1.FileName <> SaveFileDialog1.FileName) Then
                    Me.Cursor = Cursors.Wait
                    Dim xDoc As New System.Xml.XmlDocument
                    xDoc.Load(OpenFileDialog1.FileName)
                    Dim oWriter As System.Xml.XmlTextWriter = New System.Xml.XmlTextWriter(SaveFileDialog1.FileName, System.Text.Encoding.UTF8)
                    oWriter.IndentChar = ControlChars.Tab
                    oWriter.Formatting = System.Xml.Formatting.Indented
                    xDoc.Save(oWriter)
                    oWriter.Close()
                    Me.Cursor = Cursors.Arrow
                End If
            End If
        End If
    End Sub

    Private Function ValidateMSP2003(ByVal sFileName As String) As Boolean
        Dim sFile As String = g_ReadFile(sFileName)
        If sFile.Contains("<Project ") = False Then
            Return False
        End If
        If sFile.Contains("<SaveVersion>") = True Then
            Return False
        End If
        Return True
    End Function

#End Region

End Class
