Option Explicit On
Imports AGVBW

Partial Public Class fMSProject12

    Private oMP12 As MSP2007.MP12
    Private mp_lControlDraw As Integer = 0
    Private mp_lControlRedrawn As Integer = 0
    Private Const mp_sFontName As String = "Tahoma"

#Region "Constructors"

#End Region

#Region "Form Loaded"

    Private Sub Window1_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        Me.Title = "The Source Code Store - ActiveGantt Scheduler Control Version " & ActiveGanttVBWCtl1.Version & " - Microsoft Project 2007 integration using XML Files and the MSP2007 Integration Library"
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

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ScrollBar")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Color.FromArgb(255, 122, 151, 193)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 139, 144, 150)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ArrowButtons")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Color.FromArgb(255, 122, 151, 193)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 150, 158, 168)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ThumbButtonH")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 240, 240, 240)
        oStyle.EndGradientColor = Color.FromArgb(255, 165, 186, 207)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 138, 145, 153)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ThumbButtonV")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL
        oStyle.StartGradientColor = Color.FromArgb(255, 240, 240, 240)
        oStyle.EndGradientColor = Color.FromArgb(255, 165, 186, 207)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 138, 145, 153)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("TimeLineTiers")
        oStyle.Font = New Font(mp_sFontName, 7, System.Windows.FontWeights.Normal)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_TRANSPARENT
        oStyle.CustomBorderStyle.Left = True
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Right = False
        oStyle.CustomBorderStyle.Bottom = True
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.BorderColor = Color.FromArgb(255, 197, 206, 216)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("TimeLine")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.StartGradientColor = Color.FromArgb(255, 179, 206, 235)
        oStyle.EndGradientColor = Color.FromArgb(255, 161, 193, 232)
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_NONE

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ColumnStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.StartGradientColor = Color.FromArgb(255, 179, 206, 235)
        oStyle.EndGradientColor = Color.FromArgb(255, 161, 193, 232)
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Left = False
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Right = True
        oStyle.CustomBorderStyle.Bottom = True
        oStyle.BorderColor = Color.FromArgb(255, 197, 206, 216)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("TaskStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.StartGradientColor = Color.FromArgb(255, 240, 240, 240)
        oStyle.EndGradientColor = Color.FromArgb(255, 0, 0, 255)
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 148, 152, 179)
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.TextPlacement = E_TEXTPLACEMENT.SCP_EXTERIORPLACEMENT
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT
        oStyle.TextXMargin = 10
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 10
        oStyle.PredecessorStyle.LineColor = Color.FromArgb(255, 160, 160, 160)
        oStyle.MilestoneStyle.ShapeIndex = GRE_FIGURETYPE.FT_DIAMOND
        oStyle.MilestoneStyle.FillColor = Colors.Blue
        oStyle.MilestoneStyle.BorderColor = Colors.Blue
        oStyle.PredecessorStyle.XOffset = 4
        oStyle.PredecessorStyle.YOffset = 4

        oStyle = ActiveGanttVBWCtl1.Styles.Add("SummaryStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.StartGradientColor = Color.FromArgb(255, 0, 0, 0)
        oStyle.EndGradientColor = Color.FromArgb(255, 240, 240, 240)
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.BackColor = Colors.Black
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Colors.Black
        oStyle.FillMode = GRE_FILLMODE.FM_UPPERHALFFILLED
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 10
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.TextPlacement = E_TEXTPLACEMENT.SCP_EXTERIORPLACEMENT
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT
        oStyle.TextXMargin = 10
        oStyle.PredecessorStyle.LineColor = Colors.Black
        oStyle.TaskStyle.StartShapeIndex = GRE_FIGURETYPE.FT_PROJECTDOWN
        oStyle.TaskStyle.EndShapeIndex = GRE_FIGURETYPE.FT_PROJECTDOWN

        oStyle = ActiveGanttVBWCtl1.Styles.Add("NodeStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Color.FromArgb(255, 197, 206, 216)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False

        oStyle = ActiveGanttVBWCtl1.Styles.Add("CellStyleKeyColumn")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Color.FromArgb(255, 197, 206, 216)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT
        oStyle.TextXMargin = 4

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ClientAreaStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_NONE

        ActiveGanttVBWCtl1.AllowRowMove = True
        ActiveGanttVBWCtl1.AllowRowSize = True
        ActiveGanttVBWCtl1.AllowAdd = False
        ActiveGanttVBWCtl1.Style.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        ActiveGanttVBWCtl1.Style.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        ActiveGanttVBWCtl1.Style.BorderColor = Color.FromArgb(255, 122, 151, 193)
        ActiveGanttVBWCtl1.Style.BorderWidth = 1

        ActiveGanttVBWCtl1.Splitter.Type = E_SPLITTERTYPE.SA_USERDEFINED
        ActiveGanttVBWCtl1.Splitter.Width = 4
        ActiveGanttVBWCtl1.Splitter.SetColor(1, Color.FromArgb(255, 197, 206, 216))
        ActiveGanttVBWCtl1.Splitter.SetColor(2, Colors.White)
        ActiveGanttVBWCtl1.Splitter.SetColor(3, Colors.White)
        ActiveGanttVBWCtl1.Splitter.SetColor(4, Color.FromArgb(255, 197, 206, 216))
        ActiveGanttVBWCtl1.Splitter.Position = 285


        ActiveGanttVBWCtl1.Treeview.Images = True
        ActiveGanttVBWCtl1.Treeview.CheckBoxes = True
        ActiveGanttVBWCtl1.Treeview.FullColumnSelect = True
        ActiveGanttVBWCtl1.Treeview.TreeLines = False

        Dim oColumn As clsColumn

        oColumn = ActiveGanttVBWCtl1.Columns.Add("ID", "", 30, "")
        oColumn.StyleIndex = "ColumnStyle"
        oColumn.AllowTextEdit = True

        oColumn = ActiveGanttVBWCtl1.Columns.Add("Task Name", "", 300, "")
        oColumn.StyleIndex = "ColumnStyle"
        oColumn.AllowTextEdit = True

        ActiveGanttVBWCtl1.TreeviewColumnIndex = 2

        ActiveGanttVBWCtl1.ScrollBarSeparator.Style.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        ActiveGanttVBWCtl1.ScrollBarSeparator.Style.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        ActiveGanttVBWCtl1.ScrollBarSeparator.Style.BackColor = Color.FromArgb(255, 164, 196, 237)

        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.TimerInterval = 50
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButtonV"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonV"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonV"

        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButtonH"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonH"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonH"

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
        oView.TimeLine.TimeLineScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButtonH"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonH"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonH"
        oView.TimeLine.StyleIndex = "TimeLine"
        oView.ClientArea.DetectConflicts = False
        oView.ClientArea.Grid.Color = Color.FromArgb(255, 197, 206, 216)

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
        oView.TimeLine.TimeLineScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButtonH"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonH"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonH"
        oView.TimeLine.StyleIndex = "TimeLine"
        oView.ClientArea.DetectConflicts = False
        oView.ClientArea.Grid.Color = Color.FromArgb(255, 197, 206, 216)

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
        oView.TimeLine.TimeLineScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButtonH"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonH"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonH"
        oView.TimeLine.StyleIndex = "TimeLine"
        oView.ClientArea.DetectConflicts = False
        oView.ClientArea.Grid.Color = Color.FromArgb(255, 197, 206, 216)

        oView = ActiveGanttVBWCtl1.Views.Add(E_INTERVAL.IL_HOUR, 3, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
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
        oView.TimeLine.TimeLineScrollBar.SmallChange = 3
        oView.TimeLine.TimeLineScrollBar.LargeChange = 960
        oView.TimeLine.TimeLineScrollBar.Max = 8000
        oView.TimeLine.TimeLineScrollBar.Value = 0
        oView.TimeLine.TimeLineScrollBar.Enabled = True
        oView.TimeLine.TimeLineScrollBar.Visible = True
        oView.TimeLine.TimeLineScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButtonH"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonH"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonH"
        oView.TimeLine.StyleIndex = "TimeLine"
        oView.ClientArea.DetectConflicts = False
        oView.ClientArea.Grid.Color = Color.FromArgb(255, 197, 206, 216)

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
        oView.TimeLine.TimeLineScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButtonH"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonH"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonH"
        oView.TimeLine.StyleIndex = "TimeLine"
        oView.ClientArea.DetectConflicts = False
        oView.ClientArea.Grid.Color = Color.FromArgb(255, 197, 206, 216)

        ActiveGanttVBWCtl1.CurrentView = "5"

    End Sub

    Private Sub AGSetStartDate(ByVal dtStart As AGVBW.DateTime)
        Dim i As Integer
        For i = 1 To ActiveGanttVBWCtl1.Views.Count
            ActiveGanttVBWCtl1.Views.Item(i).TimeLine.TimeLineScrollBar.StartDate = dtStart
        Next
    End Sub

    Private Sub MP12_To_AG()
        Dim oAGTask As clsTask
        Dim oAGRow As clsRow
        Dim oMPTask As MSP2007.Task
        Dim dtStartDate As AGVBW.DateTime = AGVBW.DateTime.Now
        Dim i As Integer
        Dim j As Integer
        '// Load Project Tasks
        For i = 1 To oMP12.oTasks.Count
            oMPTask = oMP12.oTasks.Item(i)
            oAGRow = ActiveGanttVBWCtl1.Rows.Add("K" & oMPTask.lUID.ToString())
            oAGRow.Cells.Item("1").Text = oMPTask.lUID.ToString()
            oAGRow.Cells.Item("1").StyleIndex = "CellStyleKeyColumn"
            oAGRow.Height = 20
            oAGRow.ClientAreaStyleIndex = "ClientAreaStyle"
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
            oAGRow.Node.StyleIndex = "NodeStyle"
            If oMPTask.sNotes.Length > 0 Then
                oAGRow.Node.Image = GetImage(0)
                oAGRow.Node.ImageVisible = True
            End If
        Next
        ActiveGanttVBWCtl1.Rows.UpdateTree()
        '// Indent & set Predecessors
        For i = 1 To oMP12.oTasks.Count
            oMPTask = oMP12.oTasks.Item(i)
            oAGRow = ActiveGanttVBWCtl1.Rows.Item(i)
            oAGTask = ActiveGanttVBWCtl1.Tasks.Item(i)
            If oAGRow.Node.Children > 0 Then
                oAGTask.StyleIndex = "SummaryStyle"
            Else
                oAGTask.StyleIndex = "TaskStyle"
            End If
            For j = 1 To oMPTask.oPredecessorLink_C.Count
                Dim oMPPredecessor As MSP2007.TaskPredecessorLink
                oMPPredecessor = oMPTask.oPredecessorLink_C.Item(j)
                ActiveGanttVBWCtl1.Predecessors.Add("K" & oMPTask.lUID.ToString(), "K" & oMPPredecessor.lPredecessorUID.ToString(), GetAGPredecessorType(oMPPredecessor.yType), "", "TaskStyle")
            Next
        Next
        'Assignments
        For i = 1 To oMP12.oAssignments.Count
            Dim oAssignment As MSP2007.Assignment
            oAssignment = oMP12.oAssignments.Item(i)
            oAGTask = ActiveGanttVBWCtl1.Tasks.Item("K" & oAssignment.lTaskUID)
            If oAGTask.StartDate <> oAGTask.EndDate Then
                If oAssignment.lResourceUID > 0 Then
                    If oAGTask.Text.Length = 0 Then
                        oAGTask.Text = oMP12.oResources.Item("K" & oAssignment.lResourceUID).sName
                    Else
                        oAGTask.Text = oAGTask.Text & ", " & oMP12.oResources.Item("K" & oAssignment.lResourceUID).sName
                    End If
                End If
            End If
        Next
        dtStartDate = ActiveGanttVBWCtl1.MathLib.DateTimeAdd(E_INTERVAL.IL_DAY, -3, dtStartDate)
        AGSetStartDate(dtStartDate)
    End Sub

    Private Function GetAGPredecessorType(ByVal MPPredecessorType As MSP2007.E_TYPE_5) As AGVBW.E_CONSTRAINTTYPE
        Select Case MPPredecessorType
            Case MSP2007.E_TYPE_5.T_5_FF
                Return AGVBW.E_CONSTRAINTTYPE.PCT_END_TO_END
            Case MSP2007.E_TYPE_5.T_5_FS
                Return AGVBW.E_CONSTRAINTTYPE.PCT_END_TO_START
            Case MSP2007.E_TYPE_5.T_5_SF
                Return AGVBW.E_CONSTRAINTTYPE.PCT_START_TO_END
            Case MSP2007.E_TYPE_5.T_5_SS
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

    Private Sub AG_To_MP12()

    End Sub

#End Region

#Region "ActiveGantt Event Handlers"

    Private Sub ActiveGanttVBWCtl1_CustomTierDraw(ByVal sender As System.Object, ByVal e As AGVBW.CustomTierDrawEventArgs) Handles ActiveGanttVBWCtl1.CustomTierDraw
        If e.TierPosition = E_TIERPOSITION.SP_UPPER Then
            e.StyleIndex = "TimeLineTiers"
            If System.Convert.ToInt32(ActiveGanttVBWCtl1.CurrentView) <= 4 Then
                e.Text = e.StartDate.Year & " Q" & e.StartDate.Quarter
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
        oMP12 = New MSP2007.MP12()
        OpenFileDialog1.Title = "Load MS-Project 2007 XML File"
        OpenFileDialog1.InitialDirectory = g_GetAppLocation() & "\MSP2007\"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "XML File (*.xml)|*.xml|All Files (*.*)|*.*"
        If (OpenFileDialog1.ShowDialog(Me) = True) Then
            If ValidateMSP2007(OpenFileDialog1.FileName) = False Then
                MsgBox("The file selected is not a valid Microsoft Project 2007 XML File.", MsgBoxStyle.OkOnly)
            Else
                Me.Cursor = Cursors.Wait
                ActiveGanttVBWCtl1.Clear()
                oMP12.ReadXML(OpenFileDialog1.FileName)
                Me.Cursor = Cursors.Wait
                InitializeAG()
                MP12_To_AG()
                ActiveGanttVBWCtl1.Redraw()
                ActiveGanttVBWCtl1.VerticalScrollBar.LargeChange = ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.LastVisibleRow - ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.FirstVisibleRow
                ActiveGanttVBWCtl1.Redraw()
                Me.Cursor = Cursors.Arrow
            End If
        End If
    End Sub

    Private Sub SaveXML()
        Dim SaveFileDialog1 As New Microsoft.Win32.SaveFileDialog()
        SaveFileDialog1.Title = "Save As MS-Project 2007 XML File"
        SaveFileDialog1.InitialDirectory = g_GetAppLocation() & "\MSP2007\"
        SaveFileDialog1.Filter = "XML File|*.xml"
        SaveFileDialog1.OverwritePrompt = True
        If (SaveFileDialog1.ShowDialog(Me) = True) Then
            Me.Cursor = Cursors.Wait
            AG_To_MP12()
            oMP12.WriteXML(SaveFileDialog1.FileName)
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
        OpenFileDialog1.InitialDirectory = g_GetAppLocation() & "\MSP2007\"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "XML File (*.xml)|*.xml|All Files (*.*)|*.*"
        If (OpenFileDialog1.ShowDialog(Me) = True) Then
            SaveFileDialog1.Title = "Save XML File As"
            SaveFileDialog1.InitialDirectory = g_GetAppLocation() & "\MSP2007\"
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

    Private Function ValidateMSP2007(ByVal sFileName As String) As Boolean
        Dim sFile As String = g_ReadFile(sFileName)
        If sFile.Contains("<Project ") = False Then
            Return False
        End If
        If sFile.Contains("<SaveVersion>12</SaveVersion>") = False Then
            Return False
        End If
        Return True
    End Function

#End Region

End Class
