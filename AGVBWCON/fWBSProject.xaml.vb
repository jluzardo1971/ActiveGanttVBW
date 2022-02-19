Imports AGVBW
Imports System.Data.OleDb
Imports System.Data

Partial Public Class fWBSProject

    Private mp_dtStartDate As AGVBW.DateTime
    Private mp_dtEndDate As AGVBW.DateTime
    Private Const mp_sFontName As String = "Tahoma"
    Friend mp_yDataSourceType As E_DATASOURCETYPE
    '//XML
    Friend mp_otb_GuysStThomas As DataSet
    Friend mp_otb_GuysStThomas_Predecessors As DataSet
    Private mp_bBluePercentagesVisible As Boolean = True
    Private mp_bGreenPercentagesVisible As Boolean = True
    Private mp_bRedPercentagesVisible As Boolean = True


#Region "Constructors"

    Friend Sub New(ByVal yDataSourceType As E_DATASOURCETYPE)
        InitializeComponent()
        mp_yDataSourceType = yDataSourceType
    End Sub

#End Region

#Region "Form Loaded"

    Private Sub fWBSProject_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        ActiveGanttVBWCtl1.Visibility = Windows.Visibility.Hidden
        Me.WindowState = Windows.WindowState.Maximized

        mp_dtStartDate = New AGVBW.DateTime()
        mp_dtEndDate = New AGVBW.DateTime()
        Dim dtStartDate As AGVBW.DateTime = New AGVBW.DateTime()
        Me.Title = "Work Breakdown Structure (WBS) Project Management Example - "
        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            Me.Title = Me.Title & "Microsoft Access data source (32bit compatible only) - "
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            Me.Title = Me.Title & "XML data source (32bit and 64bit compatible) - "
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            Me.Title = Me.Title & "No data source (32bit and 64bit compatible) - "
        End If
        Me.Title = Me.Title & "ActiveGanttVBW Version: " & ActiveGanttVBWCtl1.Version

        If mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            Globals.g_VerifyWriteAccess("HPM_XML")
        End If

        Dim oStyle As clsStyle = Nothing
        Dim oView As clsView = Nothing

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ControlStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.BorderColor = Color.FromArgb(255, 100, 145, 204)
        oStyle.BackColor = Color.FromArgb(255, 240, 240, 240)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ScrollBar")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 150, 150, 150)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ArrowButtons")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 150, 150, 150)

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

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ThumbButtonHP")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 165, 186, 207)
        oStyle.EndGradientColor = Color.FromArgb(255, 240, 240, 240)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 138, 145, 153)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ThumbButtonVP")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL
        oStyle.StartGradientColor = Color.FromArgb(255, 165, 186, 207)
        oStyle.EndGradientColor = Color.FromArgb(255, 240, 240, 240)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 138, 145, 153)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ColumnStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.StartGradientColor = Color.FromArgb(255, 179, 206, 235)
        oStyle.EndGradientColor = Color.FromArgb(255, 161, 193, 232)
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.CustomBorderStyle.Left = False
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Right = True
        oStyle.CustomBorderStyle.Bottom = True
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.BorderColor = Color.FromArgb(255, 100, 145, 204)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ScrollBarSeparatorStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 150, 150, 150)

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

        oStyle = ActiveGanttVBWCtl1.Styles.Add("NodeRegular")
        oStyle.Font = New Font(mp_sFontName, 8, System.Windows.FontWeights.Normal)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False

        oStyle = ActiveGanttVBWCtl1.Styles.Add("NodeRegularChecked")
        oStyle.Font = New Font(mp_sFontName, 8, System.Windows.FontWeights.Normal)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.LightSteelBlue
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False

        oStyle = ActiveGanttVBWCtl1.Styles.Add("NodeBold")
        oStyle.Font = New Font(mp_sFontName, 8, System.Windows.FontWeights.Bold)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False

        oStyle = ActiveGanttVBWCtl1.Styles.Add("NodeBoldChecked")
        oStyle.Font = New Font(mp_sFontName, 8, System.Windows.FontWeights.Bold)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.LightSteelBlue
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ClientAreaChecked")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.LightSteelBlue

        oStyle = ActiveGanttVBWCtl1.Styles.Add("NormalTask")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Color.FromArgb(255, 100, 145, 204)
        oStyle.BorderColor = Colors.Blue
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 10
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.StartGradientColor = Colors.White
        oStyle.EndGradientColor = Color.FromArgb(255, 100, 145, 204)
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.PredecessorStyle.LineColor = Color.FromArgb(255, 100, 145, 204)
        oStyle.MilestoneStyle.ShapeIndex = GRE_FIGURETYPE.FT_DIAMOND

        oStyle = ActiveGanttVBWCtl1.Styles.Add("NormalTaskWarning")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Color.FromArgb(255, 100, 145, 204)
        oStyle.BorderColor = Colors.Red
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 10
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.StartGradientColor = Colors.White
        oStyle.EndGradientColor = Color.FromArgb(255, 100, 145, 204)
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.PredecessorStyle.LineColor = Colors.Red
        oStyle.MilestoneStyle.ShapeIndex = GRE_FIGURETYPE.FT_DIAMOND

        oStyle = ActiveGanttVBWCtl1.Styles.Add("SelectedPredecessor")
        oStyle.PredecessorStyle.LineColor = Colors.Green

        oStyle = ActiveGanttVBWCtl1.Styles.Add("GreenSummary")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.Green
        oStyle.BorderColor = Colors.Green
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 10
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.StartGradientColor = Colors.White
        oStyle.EndGradientColor = Colors.Green
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.MilestoneStyle.ShapeIndex = GRE_FIGURETYPE.FT_DIAMOND
        oStyle.SelectionRectangleStyle.Visible = False
        oStyle.TaskStyle.EndFillColor = Colors.Green
        oStyle.TaskStyle.EndBorderColor = Colors.Green
        oStyle.TaskStyle.StartFillColor = Colors.Green
        oStyle.TaskStyle.StartBorderColor = Colors.Green
        oStyle.TaskStyle.StartShapeIndex = GRE_FIGURETYPE.FT_PROJECTDOWN
        oStyle.TaskStyle.EndShapeIndex = GRE_FIGURETYPE.FT_PROJECTDOWN
        oStyle.FillMode = GRE_FILLMODE.FM_UPPERHALFFILLED

        oStyle = ActiveGanttVBWCtl1.Styles.Add("RedSummary")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.Red
        oStyle.BorderColor = Colors.Red
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 10
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.StartGradientColor = Colors.White
        oStyle.EndGradientColor = Colors.Red
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.MilestoneStyle.ShapeIndex = GRE_FIGURETYPE.FT_DIAMOND
        oStyle.SelectionRectangleStyle.Visible = False
        oStyle.TaskStyle.EndFillColor = Colors.Red
        oStyle.TaskStyle.EndBorderColor = Colors.Red
        oStyle.TaskStyle.StartFillColor = Colors.Red
        oStyle.TaskStyle.StartBorderColor = Colors.Red
        oStyle.TaskStyle.StartShapeIndex = GRE_FIGURETYPE.FT_PROJECTDOWN
        oStyle.TaskStyle.EndShapeIndex = GRE_FIGURETYPE.FT_PROJECTDOWN
        oStyle.FillMode = GRE_FILLMODE.FM_UPPERHALFFILLED

        oStyle = ActiveGanttVBWCtl1.Styles.Add("BluePercentages")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.Blue
        oStyle.BorderColor = Colors.Blue
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.OffsetTop = 8
        oStyle.OffsetBottom = 4
        oStyle.SelectionRectangleStyle.Visible = True
        oStyle.TextVisible = False
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID

        oStyle = ActiveGanttVBWCtl1.Styles.Add("GreenPercentages")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.Green
        oStyle.BorderColor = Colors.Green
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 5
        oStyle.SelectionRectangleStyle.Visible = False
        oStyle.TextVisible = False
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID

        oStyle = ActiveGanttVBWCtl1.Styles.Add("RedPercentages")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.Red
        oStyle.BorderColor = Colors.Red
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 5
        oStyle.SelectionRectangleStyle.Visible = False
        oStyle.TextVisible = False
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID

        oStyle = ActiveGanttVBWCtl1.Styles.Add("InvisiblePercentages")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 5
        oStyle.SelectionRectangleStyle.Visible = False
        oStyle.TextVisible = False
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ClientAreaStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Color.FromArgb(255, 197, 206, 216)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False
        oStyle.CustomBorderStyle.Right = False

        oStyle = ActiveGanttVBWCtl1.Styles.Add("CellStyleKeyColumn")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT
        oStyle.TextXMargin = 4

        oStyle = ActiveGanttVBWCtl1.Styles.Add("CellStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False

        oStyle = ActiveGanttVBWCtl1.Styles.Add("CellStyleKeyColumnChecked")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackColor = Colors.LightSteelBlue
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT
        oStyle.TextXMargin = 4

        oStyle = ActiveGanttVBWCtl1.Styles.Add("CellStyleChecked")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackColor = Colors.LightSteelBlue
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False

        ActiveGanttVBWCtl1.ControlTag = "WBSProject"
        ActiveGanttVBWCtl1.StyleIndex = "ControlStyle"
        ActiveGanttVBWCtl1.ScrollBarSeparator.StyleIndex = "ScrollBarSeparatorStyle"
        ActiveGanttVBWCtl1.AllowRowMove = True
        ActiveGanttVBWCtl1.AllowRowSize = True
        ActiveGanttVBWCtl1.AddMode = E_ADDMODE.AT_BOTH

        Dim oColumn As clsColumn

        oColumn = ActiveGanttVBWCtl1.Columns.Add("ID", "", 30, "")
        oColumn.StyleIndex = "ColumnStyle"
        oColumn.AllowTextEdit = True

        oColumn = ActiveGanttVBWCtl1.Columns.Add("Task Name", "", 300, "")
        oColumn.StyleIndex = "ColumnStyle"
        oColumn.AllowTextEdit = True

        oColumn = ActiveGanttVBWCtl1.Columns.Add("StartDate", "", 125, "")
        oColumn.StyleIndex = "ColumnStyle"
        oColumn.AllowTextEdit = True

        oColumn = ActiveGanttVBWCtl1.Columns.Add("EndDate", "", 125, "")
        oColumn.StyleIndex = "ColumnStyle"
        oColumn.AllowTextEdit = True

        ActiveGanttVBWCtl1.TreeviewColumnIndex = 2

        ActiveGanttVBWCtl1.Treeview.Images = True
        ActiveGanttVBWCtl1.Treeview.CheckBoxes = True
        ActiveGanttVBWCtl1.Treeview.FullColumnSelect = True
        ActiveGanttVBWCtl1.Treeview.PlusMinusBorderColor = Color.FromArgb(255, 100, 145, 204)
        ActiveGanttVBWCtl1.Treeview.PlusMinusSignColor = Color.FromArgb(255, 100, 145, 204)
        ActiveGanttVBWCtl1.Treeview.CheckBoxBorderColor = Color.FromArgb(255, 100, 145, 204)
        ActiveGanttVBWCtl1.Treeview.TreeLineColor = Color.FromArgb(255, 100, 145, 204)

        ActiveGanttVBWCtl1.Splitter.Type = E_SPLITTERTYPE.SA_USERDEFINED
        ActiveGanttVBWCtl1.Splitter.Width = 1
        ActiveGanttVBWCtl1.Splitter.SetColor(1, Color.FromArgb(255, 100, 145, 204))
        ActiveGanttVBWCtl1.Splitter.Position = 255

        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.TimerInterval = 50
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButtonV"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonVP"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonV"

        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButtonH"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonHP"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonH"

        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            Access_LoadTasks()
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            XML_LoadTasks()
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            NoDataSource_LoadTasks()
        End If
        ActiveGanttVBWCtl1.Rows.UpdateTree()

        '// Start one month before the first task:
        dtStartDate = ActiveGanttVBWCtl1.MathLib.DateTimeAdd(E_INTERVAL.IL_MONTH, -1, mp_dtStartDate)

        oView = ActiveGanttVBWCtl1.Views.Add(E_INTERVAL.IL_HOUR, 24, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_QUARTER
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.LowerTier.Factor = 1
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TimeLineScrollBar.StartDate = dtStartDate
        oView.TimeLine.TimeLineScrollBar.Enabled = True
        oView.TimeLine.TimeLineScrollBar.Visible = False
        oView.TimeLine.TimeLineScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButtonH"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonHP"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonH"
        oView.TimeLine.StyleIndex = "TimeLine"
        oView.ClientArea.DetectConflicts = False

        oView = ActiveGanttVBWCtl1.Views.Add(E_INTERVAL.IL_HOUR, 12, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_QUARTER
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.LowerTier.Factor = 1
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TimeLineScrollBar.StartDate = dtStartDate
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
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonHP"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonH"
        oView.TimeLine.StyleIndex = "TimeLine"
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
        oView.TimeLine.TimeLineScrollBar.StartDate = dtStartDate
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
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonHP"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonH"
        oView.TimeLine.StyleIndex = "TimeLine"
        oView.ClientArea.DetectConflicts = False

        ActiveGanttVBWCtl1.CurrentView = "2"


        ActiveGanttVBWCtl1.Redraw()

        ActiveGanttVBWCtl1.Visibility = Windows.Visibility.Visible

    End Sub

#End Region

#Region "Form Resizing"

    Private Sub ResizeAG()
        If Me.WindowState = Windows.WindowState.Normal Or Me.WindowState = Windows.WindowState.Maximized Then
            ActiveGanttVBWCtl1.Width = AGContainerGrid.ActualWidth
            ActiveGanttVBWCtl1.Height = AGContainerGrid.ActualHeight
        End If
    End Sub

    Private Sub Window1_SizeChanged(ByVal sender As System.Object, ByVal e As System.Windows.SizeChangedEventArgs) Handles MyBase.SizeChanged
        ResizeAG()
    End Sub

    Private Sub fWBSProject_StateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.StateChanged
        ResizeAG()
    End Sub

#End Region

#Region "ActiveGantt Event Handlers"

    Private Sub ActiveGanttVBWCtl1_CustomTierDraw(ByVal sender As System.Object, ByVal e As CustomTierDrawEventArgs) Handles ActiveGanttVBWCtl1.CustomTierDraw
        If e.TierPosition = E_TIERPOSITION.SP_LOWER Then
            e.StyleIndex = "TimeLineTiers"
            e.Text = e.StartDate.ToString("MMM")
        ElseIf e.TierPosition = E_TIERPOSITION.SP_UPPER Then
            e.StyleIndex = "TimeLineTiers"
            e.Text = e.StartDate.Year & " Q" & e.StartDate.Quarter
        End If
    End Sub

    Private Sub ActiveGanttVBWCtl1_NodeChecked(ByVal sender As System.Object, ByVal e As NodeEventArgs) Handles ActiveGanttVBWCtl1.NodeChecked
        Dim oRow As clsRow
        oRow = ActiveGanttVBWCtl1.Rows.Item(e.Index.ToString())
        If oRow.Node.Checked = True Then
            oRow.ClientAreaStyleIndex = "ClientAreaChecked"
            oRow.Cells.Item("1").StyleIndex = "CellStyleKeyColumnChecked"
            oRow.Cells.Item("3").StyleIndex = "CellStyleChecked"
            oRow.Cells.Item("4").StyleIndex = "CellStyleChecked"
            If oRow.Node.StyleIndex = "NodeBold" Then
                oRow.Node.StyleIndex = "NodeBoldChecked"
            Else
                oRow.Node.StyleIndex = "NodeRegularChecked"
            End If
        Else
            oRow.ClientAreaStyleIndex = "ClientAreaStyle"
            oRow.Cells.Item("1").StyleIndex = "CellStyleKeyColumn"
            oRow.Cells.Item("3").StyleIndex = "CellStyle"
            oRow.Cells.Item("4").StyleIndex = "CellStyle"
            If oRow.Node.StyleIndex = "NodeBoldChecked" Then
                oRow.Node.StyleIndex = "NodeBold"
            Else
                oRow.Node.StyleIndex = "NodeRegular"
            End If
        End If
    End Sub

    Private Sub ActiveGanttVBWCtl1_ControlMouseDown(ByVal sender As System.Object, ByVal e As MouseEventArgs) Handles ActiveGanttVBWCtl1.ControlMouseDown
        If (e.EventTarget = E_EVENTTARGET.EVT_TASK Or e.EventTarget = E_EVENTTARGET.EVT_SELECTEDTASK) And e.Button = E_MOUSEBUTTONS.BTN_RIGHT Then
            Dim oForm As New fWBSProjectTaskView(Me, ActiveGanttVBWCtl1.MathLib.GetTaskIndexByPosition(e.X, e.Y))
            oForm.ShowDialog()
            e.Cancel = True
        End If
    End Sub

    Private Sub ActiveGanttVBWCtl1_ObjectAdded(ByVal sender As System.Object, ByVal e As ObjectAddedEventArgs) Handles ActiveGanttVBWCtl1.ObjectAdded
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_TASK, E_EVENTTARGET.EVT_MILESTONE
                Dim oTask As clsTask
                oTask = GetTaskByRowKey(ActiveGanttVBWCtl1.Tasks.Item(e.TaskIndex).RowKey)
                oTask.StartDate = ActiveGanttVBWCtl1.Tasks.Item(e.TaskIndex).StartDate
                oTask.EndDate = ActiveGanttVBWCtl1.Tasks.Item(e.TaskIndex).EndDate
                UpdateTask(oTask.Index)
                ActiveGanttVBWCtl1.Tasks.Remove(e.TaskIndex)
            Case E_EVENTTARGET.EVT_PREDECESSOR
                ActiveGanttVBWCtl1.Predecessors.Item(e.PredecessorObjectIndex.ToString()).StyleIndex = "NormalTask"
                ActiveGanttVBWCtl1.Predecessors.Item(e.PredecessorObjectIndex.ToString()).WarningStyleIndex = "NormalTaskWarning"
                ActiveGanttVBWCtl1.Predecessors.Item(e.PredecessorObjectIndex.ToString()).SelectedStyleIndex = "SelectedPredecessor"
                InsertPredecessor(e.PredecessorTaskKey, e.TaskKey, e.PredecessorType)
        End Select
    End Sub

    Private Sub ActiveGanttVBWCtl1_CompleteObjectMove(ByVal sender As System.Object, ByVal e As ObjectStateChangedEventArgs) Handles ActiveGanttVBWCtl1.CompleteObjectMove
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_TASK
                UpdateTask(e.Index)
        End Select
    End Sub

    Private Sub ActiveGanttVBWCtl1_CompleteObjectSize(ByVal sender As System.Object, ByVal e As ObjectStateChangedEventArgs) Handles ActiveGanttVBWCtl1.CompleteObjectSize
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_TASK
                UpdateTask(e.Index)
            Case E_EVENTTARGET.EVT_PERCENTAGE
                Dim lTaskIndex As Integer
                lTaskIndex = ActiveGanttVBWCtl1.Tasks.Item(ActiveGanttVBWCtl1.Percentages.Item(e.Index).TaskKey).Index
                UpdateTask(lTaskIndex)
        End Select
    End Sub

    Private Sub ActiveGanttVBWCtl1_ToolTipOnMouseHover(ByVal sender As Object, ByVal e As ToolTipEventArgs) Handles ActiveGanttVBWCtl1.ToolTipOnMouseHover
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_TASK, E_EVENTTARGET.EVT_SELECTEDTASK, E_EVENTTARGET.EVT_PERCENTAGE
                TaskToolTipCalculateDim(e)
                Return
        End Select
        ActiveGanttVBWCtl1.ControlToolTip.Visible = False
    End Sub

    Private Sub ActiveGanttVBWCtl1_OnMouseHoverToolTipDraw(ByVal sender As Object, ByVal e As ToolTipEventArgs) Handles ActiveGanttVBWCtl1.OnMouseHoverToolTipDraw
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_TASK, E_EVENTTARGET.EVT_SELECTEDTASK, E_EVENTTARGET.EVT_PERCENTAGE, E_EVENTTARGET.EVT_SELECTEDPERCENTAGE
                TaskToolTipDraw(e)
                e.CustomDraw = True
                Return
        End Select
    End Sub

    Private Sub ActiveGanttVBWCtl1_ToolTipOnMouseMove(ByVal sender As Object, ByVal e As ToolTipEventArgs) Handles ActiveGanttVBWCtl1.ToolTipOnMouseMove
        Select Case e.Operation
            Case E_OPERATION.EO_PERCENTAGESIZING, E_OPERATION.EO_TASKMOVEMENT, E_OPERATION.EO_TASKSTRETCHLEFT, E_OPERATION.EO_TASKSTRETCHRIGHT
                TaskToolTipCalculateDim(e)
                Return
        End Select
        ActiveGanttVBWCtl1.ControlToolTip.Visible = False
    End Sub

    Private Sub ActiveGanttVBWCtl1_OnMouseMoveToolTipDraw(ByVal sender As Object, ByVal e As ToolTipEventArgs) Handles ActiveGanttVBWCtl1.OnMouseMoveToolTipDraw
        Select Case e.Operation
            Case E_OPERATION.EO_PERCENTAGESIZING, E_OPERATION.EO_TASKMOVEMENT, E_OPERATION.EO_TASKSTRETCHLEFT, E_OPERATION.EO_TASKSTRETCHRIGHT
                TaskToolTipDraw(e)
                e.CustomDraw = True
                Return
        End Select
    End Sub

    Private Sub ActiveGanttVBWCtl1_ControlMouseWheel(ByVal sender As Object, ByVal e As AGVBW.MouseWheelEventArgs) Handles ActiveGanttVBWCtl1.ControlMouseWheel
        If (e.Delta = 0) Or (ActiveGanttVBWCtl1.VerticalScrollBar.Visible = False) Then
            Return
        End If
        Dim lDelta As Integer = System.Convert.ToInt32(-(e.Delta / 100))
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

    Private Sub ActiveGanttVBWCtl1_EndTextEdit(sender As System.Object, e As AGVBW.TextEditEventArgs) Handles ActiveGanttVBWCtl1.EndTextEdit
        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            If e.ObjectType = E_TEXTOBJECTTYPE.TOT_NODE Then
                Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
                    Dim oRow As clsRow
                    Dim sRowKey As String
                    oRow = ActiveGanttVBWCtl1.Rows.Item(e.ObjectIndex)
                    sRowKey = oRow.Key
                    sRowKey = sRowKey.Replace("K", "")
                    Dim oCmd As OleDbCommand = Nothing
                    Dim sSQL As String = "UPDATE tb_GuysStThomas SET Description='" & e.Text & "' WHERE ID = " & sRowKey
                    oConn.Open()
                    oCmd = New OleDbCommand(sSQL, oConn)
                    oCmd.ExecuteNonQuery()
                    oConn.Close()
                End Using
            End If
        End If
    End Sub

#End Region

#Region "Tooltips"

    Private Sub TaskToolTipCalculateDim(ByVal e As ToolTipEventArgs)
        Dim Index As Integer = ActiveGanttVBWCtl1.MathLib.GetTaskIndexByPosition(e.X, e.Y)
        Dim oToolTip As clsToolTip = ActiveGanttVBWCtl1.ControlToolTip
        Dim sRowKey As String
        If Index = -1 Then
            sRowKey = ActiveGanttVBWCtl1.Rows.Item(ActiveGanttVBWCtl1.MathLib.GetRowIndexByPosition(e.Y)).Key
        Else
            sRowKey = ActiveGanttVBWCtl1.Tasks.Item(Index).RowKey
        End If
        Dim sRowText As String = ActiveGanttVBWCtl1.Rows.Item(sRowKey).Text
        Dim oTypeFace As New Typeface(ActiveGanttVBWCtl1.ControlToolTip.Font.FamilyName)
        Dim oFormattedText As New FormattedText(sRowText, ActiveGanttVBWCtl1.Culture, FlowDirection.LeftToRight, oTypeFace, ActiveGanttVBWCtl1.ControlToolTip.Font.WPFFontSize, New SolidColorBrush(Colors.Black))
        oFormattedText.MaxTextWidth = 275
        oToolTip.AutomaticSizing = False
        oToolTip.Left = e.X + 20
        oToolTip.Top = e.Y - (System.Convert.ToInt32(oFormattedText.Height()) + 60) - 20
        oToolTip.Width = 300
        oToolTip.Height = System.Convert.ToInt32(oFormattedText.Height()) + 60
        If ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Width < oToolTip.Width Then
            oToolTip.Visible = False
            Return
        End If
        If ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Height < oToolTip.Height Then
            oToolTip.Visible = False
            Return
        End If
        If oToolTip.Left < ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Left Then
            oToolTip.Left = ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Left
        End If
        If oToolTip.Top < ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Top Then
            oToolTip.Top = ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Top
        End If
        If ActiveGanttVBWCtl1.ControlToolTip.Left + ActiveGanttVBWCtl1.ControlToolTip.Width > ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Right Then
            ActiveGanttVBWCtl1.ControlToolTip.Left = ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Right - oToolTip.Width
        End If
        If oToolTip.Top + oToolTip.Height > ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Bottom Then
            oToolTip.Top = ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Bottom - oToolTip.Height
        End If
        oToolTip.Visible = True
    End Sub

    Private Sub TaskToolTipDraw(ByVal e As ToolTipEventArgs)
        Dim Index As Integer
        Dim sRowKey As String
        Dim sTaskKey As String
        Dim dtStartDate As AGVBW.DateTime
        Dim dtEndDate As AGVBW.DateTime
        Dim fPercentage As Single
        Dim oPercentage As clsPercentage = Nothing
        Dim oTask As clsTask
        If e.ToolTipType = E_TOOLTIPTYPE.TPT_HOVER Then
            Index = ActiveGanttVBWCtl1.MathLib.GetTaskIndexByPosition(e.X, e.Y)
            If Index < 1 Then
                Return
            End If
            oTask = ActiveGanttVBWCtl1.Tasks.Item(Index)
            sRowKey = oTask.RowKey
            dtStartDate = oTask.StartDate
            dtEndDate = oTask.EndDate
            sTaskKey = oTask.Key
            oPercentage = GetPercentageByTaskKey(sTaskKey)
            If Not oPercentage Is Nothing Then
                fPercentage = GetPercentageByTaskKey(sTaskKey).Percent * 100
            End If
        Else
            Index = e.TaskIndex
            If e.Operation = E_OPERATION.EO_TASKMOVEMENT Then
                sRowKey = ActiveGanttVBWCtl1.Rows.Item(e.InitialRowIndex).Key
            Else
                sRowKey = ActiveGanttVBWCtl1.Rows.Item(e.RowIndex).Key
            End If
            dtStartDate = e.StartDate
            dtEndDate = e.EndDate
            sTaskKey = ActiveGanttVBWCtl1.Tasks.Item(Index).Key
            If e.Operation = E_OPERATION.EO_PERCENTAGESIZING Then
                fPercentage = (e.X - e.XStart) / (e.XEnd - e.XStart) * 100
            Else
                If Not oPercentage Is Nothing Then
                    fPercentage = oPercentage.Percent * 100
                End If
            End If
        End If
        Dim sStartDate As String = dtStartDate.ToString("ddd MMM d, yyyy")
        Dim sEndDate As String = dtEndDate.ToString("ddd MMM d, yyyy")
        Dim sFrom As String = "From: " & sStartDate & " To " & sEndDate
        Dim sDuration As String = "Duration: " & ActiveGanttVBWCtl1.MathLib.DateTimeDiff(E_INTERVAL.IL_HOUR, dtStartDate, dtEndDate) & " days"
        Dim sRowText As String = ActiveGanttVBWCtl1.Rows.Item(sRowKey).Text
        Dim sPercentage As String = Format(fPercentage, "00.00")

        Dim oImage As New Image()
        oImage.Source = ActiveGanttVBWCtl1.Rows.Item(sRowKey).Node.Image.Source
        oImage.SetValue(Canvas.LeftProperty, CDbl(3))
        oImage.SetValue(Canvas.TopProperty, CDbl(3))
        e.Graphics.Children.Add(oImage)

        Dim oTypeFace As New Typeface(ActiveGanttVBWCtl1.ControlToolTip.Font.FamilyName)
        Dim oFormattedText As New FormattedText(sRowText, ActiveGanttVBWCtl1.Culture, FlowDirection.LeftToRight, oTypeFace, ActiveGanttVBWCtl1.ControlToolTip.Font.WPFFontSize, New SolidColorBrush(Colors.Black))
        oFormattedText.MaxTextWidth = 275

        Dim oTitle As New TextBlock
        oTitle.Text = sRowText
        oTitle.FontFamily = ActiveGanttVBWCtl1.ControlToolTip.Font.GetFontFamily()
        oTitle.FontSize = ActiveGanttVBWCtl1.ControlToolTip.Font.WPFFontSize()
        oTitle.SetValue(Canvas.LeftProperty, CDbl(25))
        oTitle.SetValue(Canvas.TopProperty, CDbl(2))
        oTitle.Width = 275
        oTitle.TextWrapping = TextWrapping.Wrap
        oTitle.HorizontalAlignment = Windows.HorizontalAlignment.Center
        e.Graphics.Children.Add(oTitle)

        Dim oLine As New Line
        oLine.X1 = 0
        oLine.Y1 = oFormattedText.Height + 10
        oLine.X2 = 300
        oLine.Y2 = oFormattedText.Height + 10
        oLine.Stroke = New SolidColorBrush(Colors.Black)
        oLine.StrokeThickness = 2
        e.Graphics.Children.Add(oLine)

        Dim oDuration As New TextBlock
        oDuration.Text = sDuration
        oDuration.FontFamily = ActiveGanttVBWCtl1.ControlToolTip.Font.GetFontFamily()
        oDuration.FontSize = ActiveGanttVBWCtl1.ControlToolTip.Font.WPFFontSize()
        oDuration.SetValue(Canvas.LeftProperty, CDbl(2))
        oDuration.SetValue(Canvas.TopProperty, CDbl(oFormattedText.Height + 15))
        oDuration.Width = 300
        oDuration.TextWrapping = TextWrapping.Wrap
        oDuration.HorizontalAlignment = Windows.HorizontalAlignment.Center
        e.Graphics.Children.Add(oDuration)

        Dim oInterval As New TextBlock
        oInterval.Text = sFrom
        oInterval.FontFamily = ActiveGanttVBWCtl1.ControlToolTip.Font.GetFontFamily()
        oInterval.FontSize = ActiveGanttVBWCtl1.ControlToolTip.Font.WPFFontSize()
        oInterval.SetValue(Canvas.LeftProperty, CDbl(2))
        oInterval.SetValue(Canvas.TopProperty, CDbl((oFormattedText.Height + 15) + 15))
        oInterval.Width = 300
        oInterval.TextWrapping = TextWrapping.Wrap
        oInterval.HorizontalAlignment = Windows.HorizontalAlignment.Center
        e.Graphics.Children.Add(oInterval)

        Dim oCompleted As New TextBlock
        oCompleted.Text = "Percent Completed: " & sPercentage & "%"
        oCompleted.FontFamily = ActiveGanttVBWCtl1.ControlToolTip.Font.GetFontFamily()
        oCompleted.FontSize = ActiveGanttVBWCtl1.ControlToolTip.Font.WPFFontSize()
        oCompleted.SetValue(Canvas.LeftProperty, CDbl(2))
        oCompleted.SetValue(Canvas.TopProperty, CDbl((oFormattedText.Height + 15) + 30))
        oCompleted.Width = 300
        oCompleted.TextWrapping = TextWrapping.Wrap
        oCompleted.HorizontalAlignment = Windows.HorizontalAlignment.Center
        e.Graphics.Children.Add(oCompleted)
    End Sub

#End Region

#Region "Form Properties"

#End Region

#Region "Functions"

    Private Sub UpdateTask(ByVal Index As Integer)
        Dim oPercentage As AGVBW.clsPercentage = GetPercentageByTaskKey(ActiveGanttVBWCtl1.Tasks.Item(Index.ToString()).Key)
        Dim oTask As clsTask
        oTask = ActiveGanttVBWCtl1.Tasks.Item(Index.ToString())
        SetTaskGridColumns(oTask)
        Dim sRowKey As String = oTask.RowKey
        Dim dtStartDate As AGVBW.DateTime = oTask.StartDate
        Dim dtEndDate As AGVBW.DateTime = oTask.EndDate
        Dim oNode As clsNode = ActiveGanttVBWCtl1.Rows.Item(sRowKey).Node
        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
                Dim oCmd As OleDbCommand = Nothing
                Dim sSQL As String = "UPDATE tb_GuysStThomas SET " & _
                "StartDate = " & g_DST_ACCESS_ConvertDate(dtStartDate) & _
                ", EndDate = " & g_DST_ACCESS_ConvertDate(dtEndDate) & _
                ", PercentCompleted = " & oPercentage.Percent & _
                " WHERE ID = " & sRowKey.Replace("K", "")
                oConn.Open()
                oCmd = New OleDbCommand(sSQL, oConn)
                oCmd.ExecuteNonQuery()
                oConn.Close()
            End Using
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            Dim oDataRow As DataRow = Nothing
            oDataRow = mp_otb_GuysStThomas.Tables(1).Rows.Find(sRowKey.Replace("K", ""))
            oDataRow("StartDate") = dtStartDate.DateTimePart
            oDataRow("EndDate") = dtEndDate.DateTimePart
            oDataRow("PercentCompleted") = oPercentage.Percent
            mp_otb_GuysStThomas.WriteXml(g_GetAppLocation() & "\HPM_XML\tb_GuysStThomas.xml")
        End If
        UpdateSummary(oNode)
    End Sub

    Private Sub InsertPredecessor(ByVal PredecessorKey As String, ByVal SuccessorKey As String, ByVal PredecessorType As E_CONSTRAINTTYPE)
        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
                Dim oCmd As OleDbCommand = Nothing
                PredecessorKey = PredecessorKey.Replace("T", "")
                SuccessorKey = SuccessorKey.Replace("T", "")
                Dim sSQL As String = "INSERT INTO tb_GuysStThomas_Predecessors (lPredecessorID, lSuccessorID, yType) VALUES (" & PredecessorKey.Replace("T", "") & "," & SuccessorKey.Replace("T", "") & "," & PredecessorType & ")"
                oConn.Open()
                oCmd = New OleDbCommand(sSQL, oConn)
                oCmd.ExecuteNonQuery()
                oConn.Close()
            End Using
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            Dim oDataRow As DataRow = Nothing
            Dim oLastRow As DataRow = Nothing
            oLastRow = mp_otb_GuysStThomas_Predecessors.Tables(1).Rows(mp_otb_GuysStThomas_Predecessors.Tables(1).Rows.Count - 1)
            oDataRow = mp_otb_GuysStThomas_Predecessors.Tables(1).NewRow()
            oDataRow("lID") = DirectCast(oLastRow.Item("ID"), System.Int32) + 1
            oDataRow("lPredecessorID") = PredecessorKey.Replace("T", "")
            oDataRow("lSuccessorID") = SuccessorKey.Replace("T", "")
            oDataRow("yType") = PredecessorType
            mp_otb_GuysStThomas_Predecessors.Tables(1).Rows.Add(oDataRow)
            mp_otb_GuysStThomas_Predecessors.WriteXml(g_GetAppLocation() & "\HPM_XML\tb_GuysStThomas_Predecessors.xml")
        End If
    End Sub

    Private Sub UpdateSummary(ByRef oNode As clsNode)
        Dim oConn As OleDbConnection = Nothing
        Dim oCmd As OleDbCommand = Nothing
        Dim sSQL As String = ""
        Dim oParentNode As clsNode = Nothing
        Dim oSummaryTask As clsTask = Nothing
        Dim oSummaryPercentage As clsPercentage = Nothing
        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            oConn = New OleDbConnection(g_DST_ACCESS_GetConnectionString())
            oConn.Open()
        End If
        oParentNode = oNode.Parent
        While Not oParentNode Is Nothing
            oSummaryTask = GetTaskByRowKey(oParentNode.Row.Key)
            oSummaryPercentage = GetPercentageByTaskKey(oSummaryTask.Key)
            If Not oSummaryTask Is Nothing Then
                Dim oChildTask As clsTask = Nothing
                Dim oChildPercentage As clsPercentage = Nothing
                Dim oChildNode As clsNode = Nothing
                Dim dtSumStartDate As AGVBW.DateTime = New AGVBW.DateTime()
                Dim dtSumEndDate As AGVBW.DateTime = New AGVBW.DateTime()
                Dim lPercentagesCount As Integer = 0
                Dim fPercentagesSum As Single = 0
                Dim fPercentageAvg As Single = 0
                oChildNode = oParentNode.Child
                While Not oChildNode Is Nothing
                    oChildTask = GetTaskByRowKey(oChildNode.Row.Key)
                    oChildPercentage = GetPercentageByTaskKey(oChildTask.Key)
                    lPercentagesCount = lPercentagesCount + 1
                    fPercentagesSum = fPercentagesSum + oChildPercentage.Percent
                    If Not oChildTask Is Nothing Then
                        If dtSumStartDate.DateTimePart.Ticks() = 0 Then
                            dtSumStartDate = oChildTask.StartDate
                        Else
                            If oChildTask.StartDate < dtSumStartDate Then
                                dtSumStartDate = oChildTask.StartDate
                            End If
                        End If
                        If dtSumEndDate.DateTimePart.Ticks() = 0 Then
                            dtSumEndDate = oChildTask.EndDate
                        Else
                            If oChildTask.EndDate > dtSumEndDate Then
                                dtSumEndDate = oChildTask.EndDate
                            End If
                        End If
                    End If
                    oChildNode = oChildNode.NextSibling
                End While
                fPercentageAvg = fPercentagesSum / lPercentagesCount
                oSummaryTask.StartDate = dtSumStartDate
                oSummaryTask.EndDate = dtSumEndDate
                SetTaskGridColumns(oSummaryTask)
                oSummaryPercentage.Percent = fPercentageAvg
                If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                    sSQL = "UPDATE tb_GuysStThomas SET " & _
                    "StartDate = " & g_DST_ACCESS_ConvertDate(dtSumStartDate) & _
                    ", EndDate = " & g_DST_ACCESS_ConvertDate(dtSumEndDate) & _
                    ", PercentCompleted = " & oSummaryPercentage.Percent & _
                    " WHERE ID = " & oSummaryTask.RowKey.Replace("K", "")
                    oCmd = New OleDbCommand(sSQL, oConn)
                    oCmd.ExecuteNonQuery()
                ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                    Dim oDataRow As DataRow = Nothing
                    oDataRow = mp_otb_GuysStThomas.Tables(1).Rows.Find(oSummaryTask.RowKey.Replace("K", ""))
                    oDataRow("StartDate") = dtSumStartDate.DateTimePart
                    oDataRow("EndDate") = dtSumEndDate.DateTimePart
                    oDataRow("PercentCompleted") = oSummaryPercentage.Percent
                    mp_otb_GuysStThomas.WriteXml(g_GetAppLocation() & "\HPM_XML\tb_GuysStThomas.xml")
                End If
            End If
            oParentNode = oParentNode.Parent
        End While

        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            oConn.Close()
        End If

    End Sub

    Private Function GetTaskByRowKey(ByVal sRowKey As String) As clsTask
        Dim i As Integer
        Dim oTask As clsTask
        For i = 1 To ActiveGanttVBWCtl1.Tasks.Count
            oTask = ActiveGanttVBWCtl1.Tasks.Item(i)
            If oTask.RowKey = sRowKey Then
                Return oTask
            End If
        Next
        Return Nothing
    End Function

    Private Function GetPercentageByTaskKey(ByVal sTaskKey As String) As clsPercentage
        Dim i As Integer
        Dim oPercentage As clsPercentage
        For i = 1 To ActiveGanttVBWCtl1.Percentages.Count
            oPercentage = ActiveGanttVBWCtl1.Percentages.Item(i)
            If oPercentage.TaskKey = sTaskKey Then
                Return oPercentage
            End If
        Next
        Return Nothing
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
                oURI = New Uri("../Images/WBS/folderclosed.gif", UriKind.RelativeOrAbsolute)
            Case 1
                oURI = New Uri("../Images/WBS/folderopen.gif", UriKind.RelativeOrAbsolute)
            Case 2
                oURI = New Uri("../Images/WBS/modules.gif", UriKind.RelativeOrAbsolute)
            Case 3
                oURI = New Uri("../Images/WBS/task.gif", UriKind.RelativeOrAbsolute)
        End Select
        Return oURI
    End Function

#End Region

#Region "Toolbar Buttons"

    Private Sub cmdLoadXML_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdLoadXML.Click
        LoadXML()
    End Sub

    Private Sub cmdSaveXML_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdSaveXML.Click
        SaveXML()
    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdPrint.Click
        Print()
    End Sub

    Private Sub cmdZoomIn_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdZoomIn.Click
        If ActiveGanttVBWCtl1.CurrentView < 3 Then
            ActiveGanttVBWCtl1.CurrentView = ActiveGanttVBWCtl1.CurrentView + 1
            ActiveGanttVBWCtl1.Redraw()
        End If
    End Sub

    Private Sub cmdZoomOut_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdZoomOut.Click
        If ActiveGanttVBWCtl1.CurrentView > 1 Then
            ActiveGanttVBWCtl1.CurrentView = ActiveGanttVBWCtl1.CurrentView - 1
            ActiveGanttVBWCtl1.Redraw()
        End If
    End Sub

    Private Sub cmdBluePercentages_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdBluePercentages.Click
        Dim i As Integer
        Dim oPercentage As clsPercentage
        mp_bBluePercentagesVisible = Not mp_bBluePercentagesVisible
        For i = 1 To ActiveGanttVBWCtl1.Percentages.Count
            oPercentage = ActiveGanttVBWCtl1.Percentages.Item(i.ToString())
            If oPercentage.StyleIndex = "BluePercentages" Then
                oPercentage.Visible = mp_bBluePercentagesVisible
            End If
        Next
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub cmdGreenPercentages_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdGreenPercentages.Click
        Dim i As Integer
        Dim oPercentage As clsPercentage
        mp_bGreenPercentagesVisible = Not mp_bGreenPercentagesVisible
        For i = 1 To ActiveGanttVBWCtl1.Percentages.Count
            oPercentage = ActiveGanttVBWCtl1.Percentages.Item(i.ToString())
            If oPercentage.StyleIndex = "GreenPercentages" Then
                oPercentage.Visible = mp_bGreenPercentagesVisible
            End If
        Next
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub cmdRedPercentages_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdRedPercentages.Click
        Dim i As Integer
        Dim oPercentage As clsPercentage
        mp_bRedPercentagesVisible = Not mp_bRedPercentagesVisible
        For i = 1 To ActiveGanttVBWCtl1.Percentages.Count
            oPercentage = ActiveGanttVBWCtl1.Percentages.Item(i.ToString())
            If oPercentage.StyleIndex = "RedPercentages" Then
                oPercentage.Visible = mp_bRedPercentagesVisible
            End If
        Next
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub cmdProperties_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdProperties.Click
        Dim oForm As New fWBSPProperties(Me)
        oForm.ShowDialog()
    End Sub

    Private Sub cmdCheck_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdCheck.Click
        ActiveGanttVBWCtl1.CheckPredecessors()
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub cmdTooltip_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdTooltip.Click
        ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.ToolTipsVisible = Not ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.ToolTipsVisible
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdHelp.Click
        Me.Cursor = Cursors.Wait
        System.Diagnostics.Process.Start("http://www.sourcecodestore.com/Article.aspx?ID=17")
        Me.Cursor = Cursors.Arrow
    End Sub

#End Region

#Region "Menu Items"

    Private Sub mnuSaveXML_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuSaveXML.Click
        SaveXML()
    End Sub

    Private Sub mnuLoadXML_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuLoadXML.Click
        LoadXML()
    End Sub

    Private Sub mnuPrint_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuPrint.Click
        Print()
    End Sub

    Private Sub mnuClose_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuClose.Click
        Me.Close()
    End Sub

    Private Sub mnuCheckBoxes_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuCheckBoxes.Click
        mnuCheckBoxes.IsChecked = Not (Not mnuCheckBoxes.IsChecked)
        ActiveGanttVBWCtl1.Treeview.CheckBoxes = mnuCheckBoxes.IsChecked
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub mnuImages_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuImages.Click
        mnuImages.IsChecked = Not (Not mnuImages.IsChecked)
        ActiveGanttVBWCtl1.Treeview.Images = mnuImages.IsChecked
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub mnuPlusMinusSigns_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuPlusMinusSigns.Click
        mnuPlusMinusSigns.IsChecked = Not (Not mnuPlusMinusSigns.IsChecked)
        ActiveGanttVBWCtl1.Treeview.PlusMinusSigns = mnuPlusMinusSigns.IsChecked
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub mnuFullColumnSelect_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuFullColumnSelect.Click
        mnuFullColumnSelect.IsChecked = Not (Not mnuFullColumnSelect.IsChecked)
        ActiveGanttVBWCtl1.Treeview.FullColumnSelect = mnuFullColumnSelect.IsChecked
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub mnuTreeLines_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles mnuTreeLines.Click
        mnuTreeLines.IsChecked = Not (Not mnuTreeLines.IsChecked)
        ActiveGanttVBWCtl1.Treeview.TreeLines = mnuTreeLines.IsChecked
        ActiveGanttVBWCtl1.Redraw()
    End Sub

#End Region

#Region "Toolbar Button & Menu Item Functions"

    Private Sub Print()
        Dim oForm As New fPrintDialog(ActiveGanttVBWCtl1, New AGVBW.DateTime(2006, 8, 1, 0, 0, 0), New AGVBW.DateTime(2008, 1, 1, 0, 0, 0))
        oForm.ShowDialog()
    End Sub

    Private Sub SaveXML()
        Dim dlg As New Microsoft.Win32.SaveFileDialog()
        dlg.FileName = "AGVBW_WBSP"
        dlg.DefaultExt = ".xml"
        dlg.Filter = "XML Files (.xml)|*.xml"
        If dlg.ShowDialog() = True Then
            ActiveGanttVBWCtl1.WriteXML(dlg.FileName)
        End If
    End Sub

    Private Sub LoadXML()
        Dim oForm As New fLoadXML()
        oForm.ShowDialog()
    End Sub

#End Region

#Region "Load Data"

    Public Sub Access_LoadTasks()
        Dim oRow As clsRow = Nothing
        Dim oTask As clsTask = Nothing
        Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
            Dim oCmd As OleDbCommand = Nothing
            Dim oReader As OleDbDataReader = Nothing
            oConn.Open()
            oCmd = New OleDbCommand("SELECT * FROM tb_GuysStThomas", oConn)
            oReader = oCmd.ExecuteReader
            While oReader.Read = True
                oRow = ActiveGanttVBWCtl1.Rows.Add("K" & CType(oReader.Item("ID"), System.String), DirectCast(oReader.Item("Description"), System.String))
                oRow.Cells.Item("1").Text = CType(oReader.Item("ID"), System.String)
                oRow.Cells.Item("1").StyleIndex = "CellStyle"
                oRow.Node.StyleIndex = "CellStyle"
                oRow.Cells.Item("3").StyleIndex = "CellStyle"
                oRow.Cells.Item("4").StyleIndex = "CellStyle"
                oRow.Height = 20
                oRow.ClientAreaStyleIndex = "ClientAreaStyle"
                oRow.Node.AllowTextEdit = True
                If DirectCast(oReader.Item("TaskType"), System.String) = "F" Then
                    If DirectCast(oReader.Item("Depth"), System.Int32) = 0 Then
                        oRow.Node.Image = GetImage(0)
                        oRow.Node.ExpandedImage = GetImage(1)
                        oRow.Node.StyleIndex = "NodeBold"
                    Else
                        oRow.Node.Image = GetImage(2)
                        oRow.Node.StyleIndex = "NodeRegular"
                    End If
                ElseIf DirectCast(oReader.Item("TaskType"), System.String) = "A" Then
                    oRow.Node.StyleIndex = "NodeRegular"
                    oRow.Node.Image = GetImage(3)
                    oRow.Node.CheckBoxVisible = True
                End If
                oRow.Node.Depth = DirectCast(oReader.Item("Depth"), System.Int32)
                oRow.Node.ImageVisible = True
                oRow.Node.AllowTextEdit = True
                If (Not IsDBNull(oReader.Item("StartDate")) And Not IsDBNull(oReader.Item("EndDate"))) Then
                    If (mp_dtStartDate.DateTimePart.Ticks() = 0) Then
                        mp_dtStartDate = FromDate(oReader.Item("StartDate"))
                    Else
                        If (FromDate(oReader.Item("StartDate")) < mp_dtStartDate) Then
                            mp_dtStartDate = FromDate(oReader.Item("StartDate"))
                        End If
                    End If
                    If (mp_dtEndDate.DateTimePart.Ticks() = 0) Then
                        mp_dtEndDate = FromDate(oReader.Item("EndDate"))
                    Else
                        If (FromDate(oReader.Item("EndDate")) > mp_dtEndDate) Then
                            mp_dtEndDate = FromDate(oReader.Item("EndDate"))
                        End If
                    End If
                    oTask = ActiveGanttVBWCtl1.Tasks.Add("", "K" & CType(oReader.Item("ID"), System.String), FromDate(oReader.Item("StartDate")), FromDate(oReader.Item("EndDate")), "T" & CType(oReader.Item("ID"), System.String))
                    SetTaskGridColumns(oTask)
                    If DirectCast(oReader.Item("Summary"), System.Boolean) = True Then
                        '// Prevent user from moving/sizing summary tasks
                        oTask.AllowedMovement = E_MOVEMENTTYPE.MT_MOVEMENTDISABLED
                        oTask.AllowStretchLeft = False
                        oTask.AllowStretchRight = False
                        '// Prevent user from adding tasks in these Rows
                        oRow.Container = False
                        '// Apply Summary Style 
                        If oRow.Node.Depth = 0 Then
                            oTask.StyleIndex = "RedSummary"
                            ActiveGanttVBWCtl1.Percentages.Add("T" & CType(oReader.Item("ID"), System.String), "RedPercentages", DirectCast(oReader.Item("PercentCompleted"), System.Single))
                        ElseIf oRow.Node.Depth = 1 Then
                            oTask.StyleIndex = "GreenSummary"
                            ActiveGanttVBWCtl1.Percentages.Add("T" & CType(oReader.Item("ID"), System.String), "GreenPercentages", DirectCast(oReader.Item("PercentCompleted"), System.Single))
                        End If
                        ActiveGanttVBWCtl1.Percentages.Item(ActiveGanttVBWCtl1.Percentages.Count.ToString()).AllowSize = False
                    Else
                        oTask.AllowedMovement = E_MOVEMENTTYPE.MT_RESTRICTEDTOROW
                        oTask.StyleIndex = "NormalTask"
                        oTask.WarningStyleIndex = "NormalTaskWarning"
                        If DirectCast(oReader.Item("HasTasks"), System.Boolean) = False Then
                            oTask.Visible = False
                            '// Prevent user from adding tasks in these rows
                            oRow.Container = False
                            ActiveGanttVBWCtl1.Percentages.Add("T" & CType(oReader.Item("ID"), System.String), "InvisiblePercentages", DirectCast(oReader.Item("PercentCompleted"), System.Single))
                            ActiveGanttVBWCtl1.Percentages.Item(ActiveGanttVBWCtl1.Percentages.Count.ToString()).AllowSize = False
                        Else
                            ActiveGanttVBWCtl1.Percentages.Add("T" & CType(oReader.Item("ID"), System.String), "BluePercentages", DirectCast(oReader.Item("PercentCompleted"), System.Single))
                        End If
                    End If
                End If
            End While
            oReader.Close()
            oCmd = New OleDbCommand("SELECT * FROM tb_GuysStThomas_Predecessors", oConn)
            oReader = oCmd.ExecuteReader
            Do While oReader.Read
                Dim oPredecessor As clsPredecessor
                oPredecessor = ActiveGanttVBWCtl1.Predecessors.Add("T" & oReader.Item("lSuccessorID").ToString(), "T" & oReader.Item("lPredecessorID").ToString(), oReader.Item("yType"), "", "NormalTask")
                oPredecessor.LagFactor = oReader.Item("lLagFactor")
                oPredecessor.LagInterval = oReader.Item("yLagInterval")
                oPredecessor.WarningStyleIndex = "NormalTaskWarning"
                oPredecessor.SelectedStyleIndex = "SelectedPredecessor"
            Loop
            oConn.Close()
        End Using
    End Sub

    Public Sub XML_LoadTasks()
        Dim oRow As clsRow = Nothing
        Dim oTask As clsTask = Nothing
        Dim oKeys_tb_GuysStThomas(0) As DataColumn
        Dim oKeys_tb_GuysStThomas_Predecessors(0) As DataColumn


        mp_otb_GuysStThomas = New DataSet()
        mp_otb_GuysStThomas.ReadXmlSchema(g_GetAppLocation() & "\HPM_XML\tb_GuysStThomas.xsd")
        mp_otb_GuysStThomas.ReadXml(g_GetAppLocation() & "\HPM_XML\tb_GuysStThomas.xml")
        oKeys_tb_GuysStThomas(0) = mp_otb_GuysStThomas.Tables(1).Columns("ID")
        mp_otb_GuysStThomas.Tables(1).PrimaryKey = oKeys_tb_GuysStThomas

        For Each oDataRow As DataRow In mp_otb_GuysStThomas.Tables(1).Rows
            oRow = ActiveGanttVBWCtl1.Rows.Add("K" & CType(oDataRow("ID"), System.String), DirectCast(oDataRow("Description"), System.String))
            oRow.Cells.Item("1").Text = CType(oDataRow("ID"), System.String)
            oRow.Cells.Item("1").StyleIndex = "CellStyle"
            oRow.Node.StyleIndex = "CellStyle"
            oRow.Cells.Item("3").StyleIndex = "CellStyle"
            oRow.Cells.Item("4").StyleIndex = "CellStyle"
            oRow.Height = 20
            oRow.ClientAreaStyleIndex = "ClientAreaStyle"
            oRow.Node.AllowTextEdit = True
            If DirectCast(oDataRow("TaskType"), System.String) = "F" Then
                If DirectCast(oDataRow("Depth"), System.Int32) = 0 Then
                    oRow.Node.Image = GetImage(0)
                    oRow.Node.ExpandedImage = GetImage(1)
                    oRow.Node.StyleIndex = "NodeBold"
                Else
                    oRow.Node.Image = GetImage(2)
                    oRow.Node.StyleIndex = "NodeRegular"
                End If
            ElseIf DirectCast(oDataRow("TaskType"), System.String) = "A" Then
                oRow.Node.StyleIndex = "NodeRegular"
                oRow.Node.Image = GetImage(3)
                oRow.Node.CheckBoxVisible = True
            End If
            oRow.Node.Depth = DirectCast(oDataRow("Depth"), System.Int32)
            oRow.Node.ImageVisible = True

            If (Not IsDBNull(oDataRow("StartDate")) And Not IsDBNull(oDataRow("EndDate"))) Then
                If (mp_dtStartDate.DateTimePart.Ticks() = 0) Then
                    mp_dtStartDate = FromDate(oDataRow("StartDate"))
                Else
                    If (FromDate(oDataRow("StartDate")) < mp_dtStartDate) Then
                        mp_dtStartDate = FromDate(oDataRow("StartDate"))
                    End If
                End If
                If (mp_dtEndDate.DateTimePart.Ticks() = 0) Then
                    mp_dtEndDate = FromDate(oDataRow("EndDate"))
                Else
                    If (FromDate(oDataRow("EndDate")) > mp_dtEndDate) Then
                        mp_dtEndDate = FromDate(oDataRow("EndDate"))
                    End If
                End If
                oTask = ActiveGanttVBWCtl1.Tasks.Add("", "K" & CType(oDataRow("ID"), System.String), FromDate(oDataRow("StartDate")), FromDate(oDataRow("EndDate")), "T" & CType(oDataRow("ID"), System.String))
                SetTaskGridColumns(oTask)
                If DirectCast(oDataRow("Summary"), System.Boolean) = True Then
                    '// Prevent user from moving/sizing summary tasks
                    oTask.AllowedMovement = E_MOVEMENTTYPE.MT_MOVEMENTDISABLED
                    oTask.AllowStretchLeft = False
                    oTask.AllowStretchRight = False
                    '// Prevent user from adding tasks in these Rows
                    oRow.Container = False
                    '// Apply Summary Style 
                    If oRow.Node.Depth = 0 Then
                        oTask.StyleIndex = "RedSummary"
                        ActiveGanttVBWCtl1.Percentages.Add("T" & CType(oDataRow("ID"), System.String), "RedPercentages", DirectCast(oDataRow("PercentCompleted"), System.Single))
                    ElseIf oRow.Node.Depth = 1 Then
                        oTask.StyleIndex = "GreenSummary"
                        ActiveGanttVBWCtl1.Percentages.Add("T" & CType(oDataRow("ID"), System.String), "GreenPercentages", DirectCast(oDataRow("PercentCompleted"), System.Single))
                    End If
                    ActiveGanttVBWCtl1.Percentages.Item(ActiveGanttVBWCtl1.Percentages.Count.ToString()).AllowSize = False
                Else
                    oTask.AllowedMovement = E_MOVEMENTTYPE.MT_RESTRICTEDTOROW
                    oTask.StyleIndex = "NormalTask"
                    oTask.WarningStyleIndex = "NormalTaskWarning"
                    If DirectCast(oDataRow("HasTasks"), System.Boolean) = False Then
                        oTask.Visible = False
                        '// Prevent user from adding tasks in these rows
                        oRow.Container = False
                        ActiveGanttVBWCtl1.Percentages.Add("T" & CType(oDataRow("ID"), System.String), "InvisiblePercentages", DirectCast(oDataRow("PercentCompleted"), System.Single))
                        ActiveGanttVBWCtl1.Percentages.Item(ActiveGanttVBWCtl1.Percentages.Count.ToString()).AllowSize = False
                    Else
                        ActiveGanttVBWCtl1.Percentages.Add("T" & CType(oDataRow("ID"), System.String), "BluePercentages", DirectCast(oDataRow("PercentCompleted"), System.Single))
                    End If
                End If
            End If
        Next

        mp_otb_GuysStThomas_Predecessors = New DataSet()
        mp_otb_GuysStThomas_Predecessors.ReadXmlSchema(g_GetAppLocation() & "\HPM_XML\tb_GuysStThomas_Predecessors.xsd")
        mp_otb_GuysStThomas_Predecessors.ReadXml(g_GetAppLocation() & "\HPM_XML\tb_GuysStThomas_Predecessors.xml")
        oKeys_tb_GuysStThomas_Predecessors(0) = mp_otb_GuysStThomas_Predecessors.Tables(1).Columns("ID")
        mp_otb_GuysStThomas_Predecessors.Tables(1).PrimaryKey = oKeys_tb_GuysStThomas_Predecessors

        For Each oDataRow As DataRow In mp_otb_GuysStThomas_Predecessors.Tables(1).Rows
            Dim oPredecessor As clsPredecessor
            oPredecessor = ActiveGanttVBWCtl1.Predecessors.Add("T" & oDataRow("lSuccessorID").ToString(), "T" & oDataRow("lPredecessorID").ToString(), oDataRow("yType"), "", "NormalTask")
            oPredecessor.LagFactor = oDataRow("lLagFactor")
            oPredecessor.LagInterval = oDataRow("yLagInterval")
            oPredecessor.WarningStyleIndex = "NormalTaskWarning"
            oPredecessor.SelectedStyleIndex = "SelectedPredecessor"
        Next
    End Sub

    Public Sub NoDataSource_LoadTasks()
        AddRow_Task(1, 0, "A", "Capital Plan", New AGVBW.DateTime(2007, 3, 8, 12, 0, 0), New AGVBW.DateTime(2007, 10, 19, 0, 0, 0), 0.4, False, True)
        AddRow_Task(2, 0, "F", "Strategic Projects", New AGVBW.DateTime(2006, 11, 1, 12, 0, 0), New AGVBW.DateTime(2007, 9, 14, 0, 0, 0), 0.75, True, True)
        AddRow_Task(3, 1, "F", "Infrastructure Work Team", New AGVBW.DateTime(2007, 2, 1, 12, 0, 0), New AGVBW.DateTime(2007, 9, 5, 0, 0, 0), 0.77, True, True)
        AddRow_Task(4, 2, "A", "Guys Tower Façade Feasability", New AGVBW.DateTime(2007, 2, 1, 12, 0, 0), New AGVBW.DateTime(2007, 8, 1, 0, 0, 0), 0.6, False, True)
        AddRow_Task(5, 2, "A", "East Wing Cladding (inc Ward Refurbisments)", New AGVBW.DateTime(2007, 4, 21, 0, 0, 0), New AGVBW.DateTime(2007, 9, 5, 0, 0, 0), 0.94, False, True)
        AddRow_Task(6, 1, "F", "Modernisation Workstream", New AGVBW.DateTime(2007, 1, 22, 0, 0, 0), New AGVBW.DateTime(2007, 3, 27, 12, 0, 0), 0.72, True, True)
        AddRow_Task(7, 2, "A", "A&E Reconfiguration", New AGVBW.DateTime(2007, 1, 22, 0, 0, 0), New AGVBW.DateTime(2007, 3, 27, 12, 0, 0), 0.69, False, True)
        AddRow_Task(8, 2, "A", "St. Thomas Main Theatres Study", New AGVBW.DateTime(2007, 1, 28, 0, 0, 0), New AGVBW.DateTime(2007, 3, 18, 12, 0, 0), 0.75, False, True)
        AddRow_Task(9, 1, "F", "Ambulatory Workstream", New AGVBW.DateTime(2007, 3, 9, 12, 0, 0), New AGVBW.DateTime(2007, 6, 5, 12, 0, 0), 0.73, True, True)
        AddRow_Task(10, 2, "A", "PET Feasability", New AGVBW.DateTime(2007, 3, 9, 12, 0, 0), New AGVBW.DateTime(2007, 6, 5, 12, 0, 0), 0.73, False, True)
        AddRow_Task(11, 1, "F", "Cancer Workstream", New AGVBW.DateTime(2006, 11, 1, 12, 0, 0), New AGVBW.DateTime(2007, 9, 14, 0, 0, 0), 0.78, True, True)
        AddRow_Task(12, 2, "A", "Redevelopment of Guys Site Incorporating Cancer Feasability", New AGVBW.DateTime(2007, 1, 11, 0, 0, 0), New AGVBW.DateTime(2007, 8, 11, 12, 0, 0), 0.74, False, True)
        AddRow_Task(13, 2, "A", "Radiotherapy and Chemotherapy Center", New AGVBW.DateTime(2006, 11, 1, 12, 0, 0), New AGVBW.DateTime(2007, 3, 30, 12, 0, 0), 0.94, False, True)
        AddRow_Task(14, 2, "A", "Decant Facilities", New AGVBW.DateTime(2007, 5, 24, 12, 0, 0), New AGVBW.DateTime(2007, 9, 14, 0, 0, 0), 0.65, False, True)
        AddRow_Task(15, 0, "F", "Capital Projects", New AGVBW.DateTime(2006, 9, 1, 12, 0, 0), New AGVBW.DateTime(2007, 12, 12, 0, 0, 0), 0.87, True, True)
        AddRow_Task(16, 1, "A", "4th Floor Block & Refurbishment", New AGVBW.DateTime(2006, 9, 1, 12, 0, 0), New AGVBW.DateTime(2007, 2, 1, 0, 0, 0), 0.93, False, True)
        AddRow_Task(17, 1, "A", "Bio Medical Research Center & CRF", New AGVBW.DateTime(2007, 3, 2, 0, 0, 0), New AGVBW.DateTime(2007, 7, 4, 0, 0, 0), 0.91, False, True)
        AddRow_Task(18, 1, "A", "Blundell Ward Relocation Florence + Aston Key", New AGVBW.DateTime(2007, 8, 7, 12, 0, 0), New AGVBW.DateTime(2007, 11, 12, 12, 0, 0), 0.62, False, True)
        AddRow_Task(19, 1, "A", "Bostock Ward Replacement of Water Treatment Plant", New AGVBW.DateTime(2007, 3, 7, 0, 0, 0), New AGVBW.DateTime(2007, 6, 23, 12, 0, 0), 0.84, False, True)
        AddRow_Task(20, 1, "A", "Centralisation Health Record Storage", New AGVBW.DateTime(2007, 6, 22, 0, 0, 0), New AGVBW.DateTime(2007, 11, 12, 0, 0, 0), 0.78, False, True)
        AddRow_Task(21, 1, "A", "ENT & Audiology Suite Phase II", New AGVBW.DateTime(2006, 12, 31, 12, 0, 0), New AGVBW.DateTime(2007, 3, 10, 0, 0, 0), 0.75, False, True)
        AddRow_Task(22, 1, "A", "GLI Structural Monitoring & Repair", New AGVBW.DateTime(2007, 2, 12, 12, 0, 0), New AGVBW.DateTime(2007, 5, 9, 12, 0, 0), 0.91, False, True)
        AddRow_Task(23, 1, "A", "Pathology Labs (Phase 1A)", New AGVBW.DateTime(2007, 4, 2, 0, 0, 0), New AGVBW.DateTime(2007, 10, 23, 0, 0, 0), 0.95, False, True)
        AddRow_Task(24, 1, "A", "Pathology Labs (Phase 2)", New AGVBW.DateTime(2007, 1, 15, 0, 0, 0), New AGVBW.DateTime(2007, 7, 29, 12, 0, 0), 0.92, False, True)
        AddRow_Task(25, 1, "A", "Pathology: NW5 - CSR Haematology & CSR Labs", New AGVBW.DateTime(2007, 4, 9, 0, 0, 0), New AGVBW.DateTime(2007, 9, 5, 0, 0, 0), 0.88, False, True)
        AddRow_Task(26, 1, "A", "Pathology: Haematology Day Care Center Transfer (NW4 to GT4)", New AGVBW.DateTime(2006, 10, 19, 0, 0, 0), New AGVBW.DateTime(2007, 1, 12, 0, 0, 0), 0.85, False, True)
        AddRow_Task(27, 1, "A", "HDR", New AGVBW.DateTime(2007, 6, 1, 0, 0, 0), New AGVBW.DateTime(2007, 9, 3, 0, 0, 0), 0.85, False, True)
        AddRow_Task(28, 1, "A", "Kidney Treatment Center", New AGVBW.DateTime(2007, 6, 25, 0, 0, 0), New AGVBW.DateTime(2007, 11, 18, 0, 0, 0), 0.76, False, True)
        AddRow_Task(29, 1, "A", "Maternity Expansion Business Case", New AGVBW.DateTime(2006, 11, 9, 12, 0, 0), New AGVBW.DateTime(2007, 4, 6, 0, 0, 0), 0.93, False, True)
        AddRow_Task(30, 1, "A", "New Laminar Flow Theatre at Guy's", New AGVBW.DateTime(2007, 4, 25, 12, 0, 0), New AGVBW.DateTime(2007, 11, 29, 12, 0, 0), 0.89, False, True)
        AddRow_Task(31, 1, "A", "North Wing Basement Entance - Phase 2", New AGVBW.DateTime(2007, 9, 7, 0, 0, 0), New AGVBW.DateTime(2007, 11, 30, 0, 0, 0), 0.88, False, True)
        AddRow_Task(32, 1, "A", "Paediatric Neurosciences Feasibility", New AGVBW.DateTime(2006, 11, 29, 0, 0, 0), New AGVBW.DateTime(2007, 2, 10, 0, 0, 0), 0.9, False, True)
        AddRow_Task(33, 1, "A", "Fluroscopy (Imaging 2) at St. Thomas", New AGVBW.DateTime(2007, 1, 24, 0, 0, 0), New AGVBW.DateTime(2007, 6, 8, 12, 0, 0), 0.94, False, True)
        AddRow_Task(34, 1, "A", "Interventional Radiology Suite (Imaging 3) at GT3 Phase 1", New AGVBW.DateTime(2007, 6, 17, 0, 0, 0), New AGVBW.DateTime(2007, 12, 12, 0, 0, 0), 0.91, False, True)
        AddRow_Task(35, 1, "A", "Interventional Radiology Suite (Imaging 3) at GT3 Phase 2", New AGVBW.DateTime(2007, 8, 12, 0, 0, 0), New AGVBW.DateTime(2007, 12, 1, 12, 0, 0), 0.92, False, True)
        AddRow_Task(36, 1, "A", "Imaging: Radiology Environment & Waiting Areas (Imaging 2) Phases 1 & 2", New AGVBW.DateTime(2006, 11, 27, 12, 0, 0), New AGVBW.DateTime(2007, 1, 25, 12, 0, 0), 1.0, False, True)
        AddRow_Task(37, 1, "A", "Imaging: Radiology Environment & Waiting Areas (Imaging 2) Phase 3", New AGVBW.DateTime(2006, 12, 21, 0, 0, 0), New AGVBW.DateTime(2007, 1, 9, 0, 0, 0), 1.0, False, True)
        AddRow_Task(38, 1, "A", "Relocation of Pharmacy Manufacturing & QC Laboratories", New AGVBW.DateTime(2007, 6, 7, 12, 0, 0), New AGVBW.DateTime(2007, 8, 20, 12, 0, 0), 0.93, False, True)
        AddRow_Task(39, 1, "A", "Samaritan Ward - Bone marrow transplant beds", New AGVBW.DateTime(2007, 6, 1, 0, 0, 0), New AGVBW.DateTime(2007, 8, 18, 0, 0, 0), 0.94, False, True)
        AddRow_Task(40, 1, "A", "Sexual Health Relocation", New AGVBW.DateTime(2007, 1, 10, 12, 0, 0), New AGVBW.DateTime(2007, 4, 12, 12, 0, 0), 1.0, False, True)
        AddRow_Task(41, 1, "A", "St. Thomas HV Upgrade", New AGVBW.DateTime(2007, 5, 2, 12, 0, 0), New AGVBW.DateTime(2007, 6, 20, 12, 0, 0), 0.52, False, True)
        AddRow_Task(42, 1, "A", "Ultrasound (Imaging 2) at Guy's", New AGVBW.DateTime(2007, 6, 5, 12, 0, 0), New AGVBW.DateTime(2007, 6, 22, 12, 0, 0), 1.0, False, True)
        AddRow_Task(43, 1, "F", "New Schemes Approved in Year", New AGVBW.DateTime(2006, 11, 15, 12, 0, 0), New AGVBW.DateTime(2007, 9, 4, 12, 0, 0), 0.78, True, True)
        AddRow_Task(44, 2, "A", "Modular Theatres", New AGVBW.DateTime(2006, 11, 15, 12, 0, 0), New AGVBW.DateTime(2007, 1, 1, 12, 0, 0), 0.84, False, True)
        AddRow_Task(45, 2, "A", "ECH - Theatre Ventilation", New AGVBW.DateTime(2006, 12, 24, 0, 0, 0), New AGVBW.DateTime(2007, 9, 4, 12, 0, 0), 0.77, False, True)
        AddRow_Task(46, 2, "A", "Modular Pharmacy Aseptic Unit", New AGVBW.DateTime(2006, 12, 22, 12, 0, 0), New AGVBW.DateTime(2007, 1, 28, 12, 0, 0), 0.82, False, True)
        AddRow_Task(47, 2, "A", "Acute Stroke Unit Bid", New AGVBW.DateTime(2007, 4, 11, 0, 0, 0), New AGVBW.DateTime(2007, 7, 20, 0, 0, 0), 0.74, False, True)
        AddRow_Task(48, 2, "A", "Chemo Centralisation", New AGVBW.DateTime(2006, 12, 26, 0, 0, 0), New AGVBW.DateTime(2007, 3, 30, 0, 0, 0), 0.9, False, True)
        AddRow_Task(49, 2, "A", "Feasability of MRI at Guy's", New AGVBW.DateTime(2007, 5, 12, 0, 0, 0), New AGVBW.DateTime(2007, 7, 25, 0, 0, 0), 0.59, False, True)
        AddRow_Task(50, 0, "F", "Engineering", New AGVBW.DateTime(2006, 10, 17, 0, 0, 0), New AGVBW.DateTime(2007, 9, 15, 12, 0, 0), 0.7, True, True)
        AddRow_Task(51, 1, "A", "Borough Wing Theatre Ductwork and Heater Batteries", New AGVBW.DateTime(2007, 5, 2, 0, 0, 0), New AGVBW.DateTime(2007, 6, 20, 0, 0, 0), 0.85, False, True)
        AddRow_Task(52, 1, "A", "Combined Heat and Power System at Guy's", New AGVBW.DateTime(2007, 1, 20, 12, 0, 0), New AGVBW.DateTime(2007, 4, 15, 12, 0, 0), 0.88, False, True)
        AddRow_Task(53, 1, "A", "Combined Heat and Power System at St. Thomas", New AGVBW.DateTime(2007, 3, 10, 12, 0, 0), New AGVBW.DateTime(2007, 9, 15, 12, 0, 0), 0.74, False, True)
        AddRow_Task(54, 1, "A", "Electrical Power Monitoring", New AGVBW.DateTime(2006, 11, 20, 0, 0, 0), New AGVBW.DateTime(2007, 8, 22, 12, 0, 0), 0.88, False, True)
        AddRow_Task(55, 1, "A", "Guy's Lifts 101-105 (Guys Tower)", New AGVBW.DateTime(2006, 12, 6, 0, 0, 0), New AGVBW.DateTime(2007, 3, 3, 0, 0, 0), 0.88, False, True)
        AddRow_Task(56, 1, "A", "Guy's Lifts 110-114 (Guys Tower)", New AGVBW.DateTime(2007, 5, 15, 12, 0, 0), New AGVBW.DateTime(2007, 7, 1, 12, 0, 0), 0.5, False, True)
        AddRow_Task(57, 1, "A", "Motor Control Panel Refurbishment", New AGVBW.DateTime(2007, 1, 9, 0, 0, 0), New AGVBW.DateTime(2007, 6, 13, 0, 0, 0), 0.7, False, True)
        AddRow_Task(58, 1, "A", "North Wing / Lambeth Wing Air Supply Plants", New AGVBW.DateTime(2007, 1, 13, 0, 0, 0), New AGVBW.DateTime(2007, 4, 19, 0, 0, 0), 0.21, False, True)
        AddRow_Task(59, 1, "A", "North Wing Chiller Replacement", New AGVBW.DateTime(2007, 1, 9, 0, 0, 0), New AGVBW.DateTime(2007, 6, 16, 0, 0, 0), 0.5, False, True)
        AddRow_Task(60, 1, "A", "North Wing Replacement Generator", New AGVBW.DateTime(2006, 12, 10, 12, 0, 0), New AGVBW.DateTime(2007, 6, 11, 0, 0, 0), 0.76, False, True)
        AddRow_Task(61, 1, "A", "NW/LW Riser Refurbishment", New AGVBW.DateTime(2007, 1, 20, 12, 0, 0), New AGVBW.DateTime(2007, 3, 17, 12, 0, 0), 0.5, False, True)
        AddRow_Task(62, 1, "A", "Satchwell BMS Upgrade", New AGVBW.DateTime(2006, 12, 16, 12, 0, 0), New AGVBW.DateTime(2007, 7, 18, 12, 0, 0), 0.91, False, True)
        AddRow_Task(63, 1, "A", "St. Thomas Increase Standby Capacity - Phase 2", New AGVBW.DateTime(2007, 1, 2, 0, 0, 0), New AGVBW.DateTime(2007, 6, 18, 0, 0, 0), 0.8, False, True)
        AddRow_Task(64, 1, "A", "Substation 3 HV Works (St. Thomas)", New AGVBW.DateTime(2007, 2, 27, 0, 0, 0), New AGVBW.DateTime(2007, 8, 10, 12, 0, 0), 0.78, False, True)
        AddRow_Task(65, 1, "A", "TB Electrical Distribution", New AGVBW.DateTime(2006, 10, 17, 0, 0, 0), New AGVBW.DateTime(2007, 6, 29, 12, 0, 0), 0.73, False, True)
        AddRow_Task(66, 1, "A", "Tower Wing Dental Theatre Air Handling Unit", New AGVBW.DateTime(2006, 12, 30, 12, 0, 0), New AGVBW.DateTime(2007, 3, 24, 12, 0, 0), 0.75, False, True)
        AddRow_Task(67, 1, "A", "Tower Wing Recovery Air Handling Unit", New AGVBW.DateTime(2007, 3, 2, 0, 0, 0), New AGVBW.DateTime(2007, 8, 8, 0, 0, 0), 0.7, False, True)
        AddRow_Task(68, 1, "A", "Water Booster Pumps - Phase 1 & 2", New AGVBW.DateTime(2007, 1, 8, 12, 0, 0), New AGVBW.DateTime(2007, 6, 14, 12, 0, 0), 0.64, False, True)
        AddRow_Task(69, 1, "A", "Water Softner - Boiler House", New AGVBW.DateTime(2007, 2, 12, 12, 0, 0), New AGVBW.DateTime(2007, 7, 30, 12, 0, 0), 0.66, False, True)
        AddRow_Task(70, 1, "A", "Energy Efficiency", New AGVBW.DateTime(2007, 3, 31, 12, 0, 0), New AGVBW.DateTime(2007, 9, 4, 12, 0, 0), 0.72, False, True)
        AddRow_Task(71, 0, "F", "PEAT Plan", New AGVBW.DateTime(2006, 11, 5, 0, 0, 0), New AGVBW.DateTime(2008, 1, 21, 0, 0, 0), 0.82, True, True)
        AddRow_Task(72, 1, "A", "Hilliers Ward Refurb St. Thomas", New AGVBW.DateTime(2007, 3, 28, 0, 0, 0), New AGVBW.DateTime(2007, 5, 23, 12, 0, 0), 0.79, False, True)
        AddRow_Task(73, 1, "A", "William Gull Ward St. Thomas", New AGVBW.DateTime(2007, 3, 20, 0, 0, 0), New AGVBW.DateTime(2007, 8, 23, 0, 0, 0), 0.77, False, True)
        AddRow_Task(74, 1, "A", "Henry Ward Day Room", New AGVBW.DateTime(2007, 4, 29, 0, 0, 0), New AGVBW.DateTime(2007, 6, 1, 0, 0, 0), 0.8, False, True)
        AddRow_Task(75, 1, "A", "Sarah Swift Ward", New AGVBW.DateTime(2006, 11, 5, 0, 0, 0), New AGVBW.DateTime(2007, 2, 3, 0, 0, 0), 0.78, False, True)
        AddRow_Task(76, 1, "A", "Victoria Ward", New AGVBW.DateTime(2007, 5, 10, 12, 0, 0), New AGVBW.DateTime(2007, 7, 14, 12, 0, 0), 0.91, False, True)
        AddRow_Task(77, 1, "A", "Appointment Center Staff Toilets", New AGVBW.DateTime(2007, 1, 16, 0, 0, 0), New AGVBW.DateTime(2007, 4, 7, 12, 0, 0), 0.77, False, True)
        AddRow_Task(78, 1, "A", "Page Ward", New AGVBW.DateTime(2007, 5, 19, 12, 0, 0), New AGVBW.DateTime(2007, 7, 16, 12, 0, 0), 0.74, False, True)
        AddRow_Task(79, 1, "A", "Nightingdale Ward - Side Rooms", New AGVBW.DateTime(2007, 2, 18, 0, 0, 0), New AGVBW.DateTime(2007, 4, 28, 0, 0, 0), 0.77, False, True)
        AddRow_Task(80, 1, "A", "Luke Ward - Side Rooms", New AGVBW.DateTime(2007, 11, 14, 12, 0, 0), New AGVBW.DateTime(2007, 12, 31, 12, 0, 0), 0.8, False, True)
        AddRow_Task(81, 1, "A", "Therapies Department - Disabled Toilets", New AGVBW.DateTime(2007, 7, 31, 12, 0, 0), New AGVBW.DateTime(2007, 9, 26, 12, 0, 0), 0.81, False, True)
        AddRow_Task(82, 1, "A", "Northumberland Ward Side Rooms", New AGVBW.DateTime(2007, 4, 18, 0, 0, 0), New AGVBW.DateTime(2007, 6, 6, 0, 0, 0), 0.83, False, True)
        AddRow_Task(83, 1, "A", "General Outpatients", New AGVBW.DateTime(2007, 10, 17, 0, 0, 0), New AGVBW.DateTime(2008, 1, 21, 0, 0, 0), 0.86, False, True)
        AddRow_Task(84, 1, "A", "Rheumatology Clinic", New AGVBW.DateTime(2007, 5, 3, 0, 0, 0), New AGVBW.DateTime(2007, 5, 28, 0, 0, 0), 0.84, False, True)
        AddRow_Task(85, 1, "A", "Diabetes Clinic", New AGVBW.DateTime(2007, 1, 8, 12, 0, 0), New AGVBW.DateTime(2007, 3, 18, 12, 0, 0), 0.86, False, True)
        AddRow_Task(86, 1, "A", "ENT Clinic", New AGVBW.DateTime(2007, 4, 14, 12, 0, 0), New AGVBW.DateTime(2007, 10, 28, 12, 0, 0), 0.91, False, True)
        AddRow_Task(87, 0, "F", "Buildings Improvement Programs", New AGVBW.DateTime(2006, 10, 18, 12, 0, 0), New AGVBW.DateTime(2007, 10, 28, 0, 0, 0), 0.75, True, True)
        AddRow_Task(88, 1, "F", "Environmental Improvement Plan", New AGVBW.DateTime(2006, 10, 18, 12, 0, 0), New AGVBW.DateTime(2007, 10, 28, 0, 0, 0), 0.75, False, False)
        AddRow_Task(89, 2, "A", "Ward Improvementrs", New AGVBW.DateTime(2006, 10, 18, 12, 0, 0), New AGVBW.DateTime(2007, 10, 15, 12, 0, 0), 0.61, False, True)
        AddRow_Task(90, 2, "A", "Outpatient / Clinics", New AGVBW.DateTime(2006, 12, 29, 0, 0, 0), New AGVBW.DateTime(2007, 8, 11, 0, 0, 0), 0.74, False, True)
        AddRow_Task(91, 2, "A", "Circulation Areas", New AGVBW.DateTime(2007, 4, 14, 12, 0, 0), New AGVBW.DateTime(2007, 10, 28, 0, 0, 0), 0.74, False, True)
        AddRow_Task(92, 2, "A", "St. Thomas Main Entrance", New AGVBW.DateTime(2007, 2, 28, 0, 0, 0), New AGVBW.DateTime(2007, 6, 8, 0, 0, 0), 0.76, False, True)
        AddRow_Task(93, 2, "A", "St. Thomas Retail Mall", New AGVBW.DateTime(2007, 1, 1, 0, 0, 0), New AGVBW.DateTime(2007, 2, 6, 0, 0, 0), 0.81, False, True)
        AddRow_Task(94, 2, "A", "Guys Main Entrance Revolving Door", New AGVBW.DateTime(2007, 3, 28, 12, 0, 0), New AGVBW.DateTime(2007, 4, 25, 12, 0, 0), 0.83, False, True)

        AddPredecessor(16, 17, E_CONSTRAINTTYPE.PCT_END_TO_START, 696, E_INTERVAL.IL_HOUR)     '//End-To-Start with lag (down)
        AddPredecessor(13, 5, E_CONSTRAINTTYPE.PCT_END_TO_START, 516, E_INTERVAL.IL_HOUR)      '//End-To-Start with lag (up)
        AddPredecessor(21, 22, E_CONSTRAINTTYPE.PCT_END_TO_START, -612, E_INTERVAL.IL_HOUR)    '//End-To-Start with lead (down)
        AddPredecessor(24, 19, E_CONSTRAINTTYPE.PCT_END_TO_START, -3468, E_INTERVAL.IL_HOUR)   '//End-To-Start with lead (up)

        AddPredecessor(18, 20, E_CONSTRAINTTYPE.PCT_START_TO_END, 2316, E_INTERVAL.IL_HOUR)    '//Start-To-End with lag (down)
        AddPredecessor(29, 26, E_CONSTRAINTTYPE.PCT_START_TO_END, 1524, E_INTERVAL.IL_HOUR)    '//Start-To-End with lag (up)
        AddPredecessor(27, 32, E_CONSTRAINTTYPE.PCT_START_TO_END, -2664, E_INTERVAL.IL_HOUR)   '//Start-To-End with lead (down)
        AddPredecessor(38, 36, E_CONSTRAINTTYPE.PCT_START_TO_END, -3192, E_INTERVAL.IL_HOUR)   '//Start-To-End with lead (up)

        AddPredecessor(12, 14, E_CONSTRAINTTYPE.PCT_START_TO_START, 3204, E_INTERVAL.IL_HOUR)  '//Start-To-Start with lag (down)
        AddPredecessor(48, 47, E_CONSTRAINTTYPE.PCT_START_TO_START, 2544, E_INTERVAL.IL_HOUR)  '//Start-To-Start with lag (up)
        AddPredecessor(52, 55, E_CONSTRAINTTYPE.PCT_START_TO_START, -1092, E_INTERVAL.IL_HOUR) '//Start-To-Start with lead (down)
        AddPredecessor(56, 53, E_CONSTRAINTTYPE.PCT_START_TO_START, -1584, E_INTERVAL.IL_HOUR) '//Start-To-Start with lead (up)

        AddPredecessor(40, 41, E_CONSTRAINTTYPE.PCT_END_TO_END, 1656, E_INTERVAL.IL_HOUR)      '//End-To-End with lag (down)
        AddPredecessor(58, 57, E_CONSTRAINTTYPE.PCT_END_TO_END, 1320, E_INTERVAL.IL_HOUR)      '//End-To-End with lag (up)
        AddPredecessor(62, 63, E_CONSTRAINTTYPE.PCT_END_TO_END, -732, E_INTERVAL.IL_HOUR)      '//End-To-End with lead (down)
        AddPredecessor(67, 65, E_CONSTRAINTTYPE.PCT_END_TO_END, -948, E_INTERVAL.IL_HOUR)      '//End-To-End with lead (up)

    End Sub

    Public Sub AddPredecessor(ByVal lPredecessorID As Integer, ByVal lSuccessorID As Integer, ByVal yType As E_CONSTRAINTTYPE, ByVal lLagFactor As Integer, ByVal yLagInterval As E_INTERVAL)
        Dim oPredecessor As clsPredecessor
        oPredecessor = ActiveGanttVBWCtl1.Predecessors.Add("T" & lSuccessorID.ToString(), "T" & lPredecessorID.ToString(), yType, "", "NormalTask")
        oPredecessor.WarningStyleIndex = "NormalTaskWarning"
        oPredecessor.SelectedStyleIndex = "SelectedPredecessor"
        oPredecessor.LagFactor = lLagFactor
        oPredecessor.LagInterval = yLagInterval
    End Sub

    Public Sub AddRow_Task(ByVal lID As Integer, ByVal lDepth As Integer, ByVal sTaskType As String, ByVal sDescription As String, ByVal dtStartDate As AGVBW.DateTime, ByVal dtEndDate As AGVBW.DateTime, ByVal fPercentCompleted As Single, ByVal bSummary As Boolean, ByVal bHasTasks As Boolean)
        Dim oRow As clsRow = Nothing
        Dim oTask As clsTask = Nothing
        oRow = ActiveGanttVBWCtl1.Rows.Add("K" & lID.ToString(), sDescription)
        oRow.Cells.Item("1").Text = lID.ToString()
        oRow.Cells.Item("1").StyleIndex = "CellStyleKeyColumn"
        oRow.Node.StyleIndex = "CellStyle"
        oRow.Cells.Item("3").StyleIndex = "CellStyle"
        oRow.Cells.Item("4").StyleIndex = "CellStyle"
        oRow.Height = 20
        oRow.ClientAreaStyleIndex = "ClientAreaStyle"
        oRow.Node.AllowTextEdit = True
        If sTaskType = "F" Then
            If lDepth = 0 Then
                oRow.Node.Image = GetImage(0)
                oRow.Node.ExpandedImage = GetImage(1)
                oRow.Node.StyleIndex = "NodeBold"
            Else
                oRow.Node.Image = GetImage(2)
                oRow.Node.StyleIndex = "NodeRegular"
            End If
        ElseIf sTaskType = "A" Then
            oRow.Node.StyleIndex = "NodeRegular"
            oRow.Node.Image = GetImage(3)
            oRow.Node.CheckBoxVisible = True
        End If
        oRow.Node.Depth = lDepth
        oRow.Node.ImageVisible = True
        oRow.Node.AllowTextEdit = True
        If (mp_dtStartDate.DateTimePart.Ticks() = 0) Then
            mp_dtStartDate = dtStartDate
        Else
            If (dtStartDate < mp_dtStartDate) Then
                mp_dtStartDate = dtStartDate
            End If
        End If
        If (mp_dtEndDate.DateTimePart.Ticks() = 0) Then
            mp_dtEndDate = dtEndDate
        Else
            If (dtEndDate > mp_dtEndDate) Then
                mp_dtEndDate = dtEndDate
            End If
        End If
        oTask = ActiveGanttVBWCtl1.Tasks.Add("", "K" & lID, dtStartDate, dtEndDate, "T" & lID.ToString())
        SetTaskGridColumns(oTask)
        If bSummary = True Then
            '// Prevent user from moving/sizing summary tasks
            oTask.AllowedMovement = E_MOVEMENTTYPE.MT_MOVEMENTDISABLED
            oTask.AllowStretchLeft = False
            oTask.AllowStretchRight = False
            '// Prevent user from adding tasks in these Rows
            oRow.Container = False
            '// Apply Summary Style 
            If oRow.Node.Depth = 0 Then
                oTask.StyleIndex = "RedSummary"
                ActiveGanttVBWCtl1.Percentages.Add("T" & lID.ToString(), "RedPercentages", fPercentCompleted)
            ElseIf oRow.Node.Depth = 1 Then
                oTask.StyleIndex = "GreenSummary"
                ActiveGanttVBWCtl1.Percentages.Add("T" & lID.ToString(), "GreenPercentages", fPercentCompleted)
            End If
            ActiveGanttVBWCtl1.Percentages.Item(ActiveGanttVBWCtl1.Percentages.Count.ToString()).AllowSize = False
        Else
            oTask.AllowedMovement = E_MOVEMENTTYPE.MT_RESTRICTEDTOROW
            oTask.StyleIndex = "NormalTask"
            oTask.WarningStyleIndex = "NormalTaskWarning"
            If bHasTasks = False Then
                oTask.Visible = False
                '// Prevent user from adding tasks in these rows
                oRow.Container = False
                ActiveGanttVBWCtl1.Percentages.Add("T" & lID.ToString(), "InvisiblePercentages", fPercentCompleted)
                ActiveGanttVBWCtl1.Percentages.Item(ActiveGanttVBWCtl1.Percentages.Count.ToString()).AllowSize = False
            Else
                ActiveGanttVBWCtl1.Percentages.Add("T" & lID.ToString(), "BluePercentages", fPercentCompleted)
            End If
        End If
    End Sub

    Private Sub SetTaskGridColumns(ByVal oTask As clsTask)
        oTask.Row.Cells.Item("3").Text = oTask.StartDate.ToString("MM/dd/yyyy")
        oTask.Row.Cells.Item("4").Text = oTask.EndDate.ToString("MM/dd/yyyy")
    End Sub

#End Region

End Class
