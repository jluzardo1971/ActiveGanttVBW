Imports System.Data.OleDb
Imports AGVBW
Imports System.Data

Partial Public Class fCarRental

    Public Enum HPE_ADDMODE
        AM_RESERVATION = 0
        AM_RENTAL = 1
        AM_MAINTENANCE = 2
    End Enum

    Private mp_yAddMode As HPE_ADDMODE = HPE_ADDMODE.AM_RENTAL
    Private mp_sAddModeStyleIndex As String
    Private mp_lZoom As Integer
    Private mp_sEditRowKey As String
    Private mp_sEditTaskKey As String
    Private Const mp_sFontName As String = "Tahoma"
    Friend mp_yDataSourceType As E_DATASOURCETYPE
    '//XML & NO DATA_SOURCE
    Friend mp_otb_CR_Rows As DataSet
    Friend mp_otb_CR_Rentals As DataSet

    Friend mp_otb_CR_Car_Types As DataSet
    Friend mp_otb_CR_US_States As DataSet
    Friend mp_otb_CR_ACRISS_Codes As DataSet
    Friend mp_otb_CR_Taxes_Surcharges_Options As DataSet

#Region "Constructors"

    Friend Sub New(ByVal yDataSourceType As E_DATASOURCETYPE)
        MyBase.New()
        InitializeComponent()
        mp_yDataSourceType = yDataSourceType
    End Sub

#End Region

#Region "Form Loaded"

    Private Sub fCarRental_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        ActiveGanttVBWCtl1.Visibility = Windows.Visibility.Hidden
        Me.WindowState = Windows.WindowState.Maximized

        Me.Title = "Vehicle Rental/Fleet Control Roster Example - "
        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            Me.Title = Me.Title & "Microsoft Access data source (32bit compatible only) - "
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            Me.Title = Me.Title & "XML data source (32bit and 64bit compatible) - "
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            Me.Title = Me.Title & "No data source (32bit and 64bit compatible) - "
        End If
        Me.Title = Me.Title & "ActiveGanttVBW Version: " & ActiveGanttVBWCtl1.Version

        If mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            g_VerifyWriteAccess("CR_XML")
            XML_Load_Car_Types()
            XML_Load_US_States()
            XML_Load_ACRISS_Codes()
            XML_Load_Taxes_Surcharges_Options()
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            NoDataSource_Load_Car_Types()
            NoDataSource_Load_US_States()
            NoDataSource_Load_ACRISS_Codes()
            NoDataSource_Load_Taxes_Surcharges_Options()
        End If


        Dim oStyle As clsStyle = Nothing
        Dim oView As clsView = Nothing
        Dim oTimeBlock As clsTimeBlock = Nothing

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ScrollBar")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 150, 158, 168)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ArrowButtons")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 150, 158, 168)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("ThumbButton")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 138, 145, 153)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("SplitterStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 109, 122, 136)
        oStyle.EndGradientColor = Color.FromArgb(255, 220, 220, 220)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("Columns")
        oStyle.Font = New Font(mp_sFontName, 8, System.Windows.FontWeights.Bold)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 148, 164, 189)
        oStyle.EndGradientColor = Color.FromArgb(255, 178, 199, 228)
        oStyle.ForeColor = Colors.White
        oStyle.BorderColor = Colors.Black
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Left = False
        oStyle.CustomBorderStyle.Top = False
        oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM

        oStyle = ActiveGanttVBWCtl1.Styles.Add("TimeLine")
        oStyle.Font = New Font(mp_sFontName, 7, System.Windows.FontWeights.Normal)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 148, 164, 189)
        oStyle.EndGradientColor = Color.FromArgb(255, 178, 199, 228)
        oStyle.ForeColor = Colors.White
        oStyle.BorderColor = Colors.Black
        oStyle.CustomBorderStyle.Left = True
        oStyle.CustomBorderStyle.Top = True
        oStyle.CustomBorderStyle.Right = False
        oStyle.CustomBorderStyle.Bottom = True
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM

        oStyle = ActiveGanttVBWCtl1.Styles.Add("TimeLineVA")
        oStyle.Font = New Font(mp_sFontName, 7, System.Windows.FontWeights.Normal)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 148, 164, 189)
        oStyle.EndGradientColor = Color.FromArgb(255, 178, 199, 228)
        oStyle.ForeColor = Colors.White
        oStyle.BorderColor = Colors.Black
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.DrawTextInVisibleArea = True

        oStyle = ActiveGanttVBWCtl1.Styles.Add("Branch")
        oStyle.Font = New Font(mp_sFontName, 9, System.Windows.FontWeights.Normal)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 109, 122, 136)
        oStyle.EndGradientColor = Color.FromArgb(255, 179, 199, 229)
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
        oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP
        oStyle.TextXMargin = 5
        oStyle.TextYMargin = 5
        oStyle.ForeColor = Colors.White
        oStyle.BorderColor = Colors.Black
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.ImageAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT
        oStyle.ImageAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM
        oStyle.ImageXMargin = 5
        oStyle.ImageYMargin = 5
        oStyle.UseMask = False

        oStyle = ActiveGanttVBWCtl1.Styles.Add("BranchCA")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 109, 122, 136)
        oStyle.EndGradientColor = Color.FromArgb(255, 179, 199, 229)
        oStyle.ForeColor = Colors.White

        oStyle = ActiveGanttVBWCtl1.Styles.Add("Weekend")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL
        oStyle.StartGradientColor = Color.FromArgb(255, 133, 143, 154)
        oStyle.EndGradientColor = Color.FromArgb(255, 172, 183, 194)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("Reservation")
        oStyle.Font = New Font(mp_sFontName, 7, System.Windows.FontWeights.Normal)
        oStyle.ForeColor = Colors.White
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
        oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP
        oStyle.TextXMargin = 5
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL
        oStyle.StartGradientColor = Color.FromArgb(255, 109, 122, 136)
        oStyle.EndGradientColor = Color.FromArgb(255, 179, 199, 229)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("Rental")
        oStyle.Font = New Font(mp_sFontName, 7, System.Windows.FontWeights.Normal)
        oStyle.ForeColor = Colors.White
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
        oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP
        oStyle.TextXMargin = 5
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL
        oStyle.StartGradientColor = Color.FromArgb(255, 162, 78, 50)
        oStyle.EndGradientColor = Color.FromArgb(255, 215, 92, 54)

        oStyle = ActiveGanttVBWCtl1.Styles.Add("Maintenance")
        oStyle.Font = New Font(mp_sFontName, 7, System.Windows.FontWeights.Normal)
        oStyle.ForeColor = Colors.White
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
        oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP
        oStyle.TextXMargin = 5
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL
        oStyle.StartGradientColor = Color.FromArgb(255, 255, 77, 1)
        oStyle.EndGradientColor = Color.FromArgb(255, 244, 172, 43)

        ActiveGanttVBWCtl1.ControlTag = "CarRental"
        ActiveGanttVBWCtl1.Columns.Add("", "", 45, "Columns")
        ActiveGanttVBWCtl1.Columns.Add("", "", 95, "Columns")
        ActiveGanttVBWCtl1.Columns.Add("", "", 250, "Columns")

        ActiveGanttVBWCtl1.Splitter.Position = 340
        ActiveGanttVBWCtl1.Splitter.Type = E_SPLITTERTYPE.SA_STYLE
        ActiveGanttVBWCtl1.Splitter.Width = 6
        ActiveGanttVBWCtl1.Splitter.StyleIndex = "SplitterStyle"

        ActiveGanttVBWCtl1.ScrollBarSeparator.StyleIndex = "ScrollBar"

        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButton"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButton"
        ActiveGanttVBWCtl1.VerticalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButton"

        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButton"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButton"
        ActiveGanttVBWCtl1.HorizontalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButton"

        ActiveGanttVBWCtl1.Culture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        ActiveGanttVBWCtl1.Culture.DateTimeFormat.AMDesignator = ActiveGanttVBWCtl1.Culture.DateTimeFormat.AMDesignator.ToLower()
        ActiveGanttVBWCtl1.Culture.DateTimeFormat.PMDesignator = ActiveGanttVBWCtl1.Culture.DateTimeFormat.PMDesignator.ToLower()

        oTimeBlock = ActiveGanttVBWCtl1.TimeBlocks.Add("")
        oTimeBlock.TimeBlockType = E_TIMEBLOCKTYPE.TBT_RECURRING
        oTimeBlock.RecurringType = E_RECURRINGTYPE.RCT_WEEK
        oTimeBlock.BaseWeekDay = E_WEEKDAY.WD_FRIDAY
        oTimeBlock.BaseDate = New AGVBW.DateTime(2013, 1, 1, 0, 0, 0)
        oTimeBlock.DurationFactor = 48
        oTimeBlock.DurationInterval = E_INTERVAL.IL_HOUR
        oTimeBlock.StyleIndex = "Weekend"

        oView = ActiveGanttVBWCtl1.Views.Add(E_INTERVAL.IL_MINUTE, 30, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.MiddleTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Interval = E_INTERVAL.IL_DAY
        oView.TimeLine.TierArea.MiddleTier.Factor = 1
        oView.TimeLine.TierArea.MiddleTier.Visible = True
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_HOUR
        oView.TimeLine.TierArea.LowerTier.Factor = 12
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TickMarkArea.StyleIndex = "TimeLine"
        oView.TimeLine.Style.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oView.TimeLine.Style.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oView.TimeLine.Style.BackColor = Colors.Black
        oView.ClientArea.Grid.VerticalLines = True
        oView.ClientArea.Grid.SnapToGrid = True
        ActiveGanttVBWCtl1.CurrentView = oView.Index.ToString()
        Zoom = 5

        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            Access_LoadRowsAndTasks()
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            XML_LoadRowsAndTasks()
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            NoDataSource_LoadRowsAndTasks()
        End If
        ActiveGanttVBWCtl1.Rows.UpdateTree()

        Mode = HPE_ADDMODE.AM_RESERVATION

        ActiveGanttVBWCtl1.CurrentViewObject.TimeLine.Position(New DateTime(2009, 6, 9))


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

    Private Sub fCarRental_SizeChanged(ByVal sender As System.Object, ByVal e As System.Windows.SizeChangedEventArgs) Handles MyBase.SizeChanged
        ResizeAG()
    End Sub

    Private Sub fCarRental_StateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.StateChanged
        ResizeAG()
    End Sub

#End Region

#Region "ActiveGantt Event Handlers"

    Private Sub ActiveGanttVBWCtl1_CustomTierDraw(ByVal sender As System.Object, ByVal e As CustomTierDrawEventArgs) Handles ActiveGanttVBWCtl1.CustomTierDraw
        If e.Interval = E_INTERVAL.IL_HOUR And e.Factor = 12 Then
            e.Text = e.StartDate.ToString("tt").ToUpper()
            e.StyleIndex = "TimeLine"
        End If
        If e.Interval = E_INTERVAL.IL_MONTH And e.Factor = 1 Then
            e.Text = e.StartDate.ToString("MMMM yyyy")
            e.StyleIndex = "TimeLineVA"
        End If
        If e.Interval = E_INTERVAL.IL_DAY And e.Factor = 1 Then
            e.Text = e.StartDate.ToString("ddd d")
            e.StyleIndex = "TimeLine"
        End If
    End Sub

    Private Sub ActiveGanttVBWCtl1_ObjectAdded(ByVal sender As System.Object, ByVal e As ObjectAddedEventArgs) Handles ActiveGanttVBWCtl1.ObjectAdded
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_TASK
                Dim oTask As clsTask = Nothing
                Dim lTaskID As Integer = 0
                Dim oDataRow As DataRow = Nothing
                oTask = ActiveGanttVBWCtl1.Tasks.Item(e.TaskIndex.ToString())
                oTask.StyleIndex = mp_sAddModeStyleIndex
                oTask.Tag = mp_yAddMode.ToString()
                If Mode = HPE_ADDMODE.AM_RESERVATION Then
                    Dim oForm As New fCarRentalReservation(PRG_DIALOGMODE.DM_ADD, Me, "")
                    oForm.ShowDialog()
                ElseIf Mode = HPE_ADDMODE.AM_RENTAL Then
                    Dim oForm As New fCarRentalReservation(PRG_DIALOGMODE.DM_ADD, Me, "")
                    oForm.ShowDialog()
                ElseIf Mode = HPE_ADDMODE.AM_MAINTENANCE Then
                    If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                        Dim oDB As clsDB = Nothing
                        oDB = New clsDB()
                        oDB.AddParameter("lRowID", oTask.RowKey.Replace("K", ""), clsDB.ParamType.PT_NUMERIC)
                        oDB.AddParameter("yMode", 2, clsDB.ParamType.PT_NUMERIC)
                        oDB.AddParameter("dtPickUp", oTask.StartDate, clsDB.ParamType.PT_DATE)
                        oDB.AddParameter("dtReturn", oTask.EndDate, clsDB.ParamType.PT_DATE)
                        oTask.Key = "K" & oDB.ExecuteInsert("tb_CR_Rentals")
                        oTask.Text = "Scheduled Maintenance"
                        oTask.Tag = CType(CType(HPE_ADDMODE.AM_MAINTENANCE, System.Int32), System.String)
                    ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                        oDataRow = mp_otb_CR_Rentals.Tables(1).NewRow()
                        oDataRow("lRowID") = oTask.RowKey.Replace("K", "")
                        oDataRow("yMode") = 2
                        oDataRow("dtPickup") = oTask.StartDate.DateTimePart
                        oDataRow("dtReturn") = oTask.EndDate.DateTimePart
                        oDataRow("bGPS") = False
                        oDataRow("bFSO") = False
                        oDataRow("bLDW") = False
                        oDataRow("bPAI") = False
                        oDataRow("bPEP") = False
                        oDataRow("bALI") = False
                        lTaskID = g_DST_XML_AutoIncrementValue(mp_otb_CR_Rentals, "lTaskID")
                        oDataRow("lTaskID") = lTaskID
                        mp_otb_CR_Rentals.Tables(1).Rows.Add(oDataRow)
                        mp_otb_CR_Rentals.WriteXml(g_GetAppLocation() & "\CR_XML\tb_CR_Rentals.xml")
                        oTask.Key = "K" & lTaskID.ToString()
                        oTask.Text = "Scheduled Maintenance"
                        oTask.Tag = CType(CType(HPE_ADDMODE.AM_MAINTENANCE, System.Int32), System.String)
                    ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                        oDataRow = mp_otb_CR_Rentals.Tables(0).NewRow()
                        oDataRow("lRowID") = oTask.RowKey.Replace("K", "")
                        oDataRow("yMode") = 2
                        oDataRow("dtPickup") = oTask.StartDate.DateTimePart
                        oDataRow("dtReturn") = oTask.EndDate.DateTimePart
                        oDataRow("bGPS") = False
                        oDataRow("bFSO") = False
                        oDataRow("bLDW") = False
                        oDataRow("bPAI") = False
                        oDataRow("bPEP") = False
                        oDataRow("bALI") = False
                        lTaskID = g_DST_NONE_AutoIncrementValue(mp_otb_CR_Rentals, "lTaskID")
                        oDataRow("lTaskID") = lTaskID
                        mp_otb_CR_Rentals.Tables(0).Rows.Add(oDataRow)
                        oTask.Key = "K" & lTaskID.ToString()
                        oTask.Text = "Scheduled Maintenance"
                        oTask.Tag = CType(CType(HPE_ADDMODE.AM_MAINTENANCE, System.Int32), System.String)
                    End If
                End If
        End Select
    End Sub

    Private Sub ActiveGanttVBWCtl1_CompleteObjectMove(ByVal sender As System.Object, ByVal e As ObjectStateChangedEventArgs) Handles ActiveGanttVBWCtl1.CompleteObjectMove
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_TASK
                Dim oTask As clsTask = Nothing
                oTask = ActiveGanttVBWCtl1.Tasks.Item(e.Index.ToString())
                CalculateRate(oTask)
            Case E_EVENTTARGET.EVT_ROW
                Dim i As Integer = 0
                Dim oRow As clsRow = Nothing
                Dim oDataRow As DataRow = Nothing
                For i = 1 To ActiveGanttVBWCtl1.Rows.Count
                    oRow = ActiveGanttVBWCtl1.Rows.Item(i.ToString())
                    If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                        g_DST_ACCESS_ExecuteNonQuery("UPDATE tb_CR_Rows SET lOrder = " & i & " WHERE lRowID = " & oRow.Key.Replace("K", ""))
                    ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                        oDataRow = mp_otb_CR_Rows.Tables(1).Rows.Find(oRow.Key.Replace("K", ""))
                        oDataRow("lOrder") = i
                        mp_otb_CR_Rows.WriteXml(g_GetAppLocation() & "\CR_XML\tb_CR_Rows.xml")
                    ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                        oDataRow = mp_otb_CR_Rows.Tables(0).Rows.Find(oRow.Key.Replace("K", ""))
                        oDataRow("lOrder") = i
                    End If
                Next
        End Select
    End Sub

    Private Sub ActiveGanttVBWCtl1_CompleteObjectSize(ByVal sender As System.Object, ByVal e As ObjectStateChangedEventArgs) Handles ActiveGanttVBWCtl1.CompleteObjectSize
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_TASK
                Dim oTask As clsTask
                oTask = ActiveGanttVBWCtl1.Tasks.Item(e.Index)
                CalculateRate(oTask)
        End Select
    End Sub

    Private Sub ActiveGanttVBWCtl1_ControlMouseDown(ByVal sender As System.Object, ByVal e As MouseEventArgs) Handles ActiveGanttVBWCtl1.ControlMouseDown
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_SELECTEDROW, E_EVENTTARGET.EVT_ROW
                If e.Button = E_MOUSEBUTTONS.BTN_LEFT Then
                    Dim oRow As clsRow
                    oRow = ActiveGanttVBWCtl1.Rows.Item(ActiveGanttVBWCtl1.MathLib.GetRowIndexByPosition(e.Y))
                    If e.X > ActiveGanttVBWCtl1.Splitter.Position - 20 And e.X < ActiveGanttVBWCtl1.Splitter.Position - 5 And e.Y < oRow.Bottom - 5 And e.Y > oRow.Bottom - 20 Then
                        oRow.Node.Expanded = Not oRow.Node.Expanded
                        ActiveGanttVBWCtl1.Redraw()
                        e.Cancel = True
                    End If
                ElseIf e.Button = E_MOUSEBUTTONS.BTN_RIGHT Then
                    e.Cancel = True
                    mp_sEditRowKey = ActiveGanttVBWCtl1.Rows.Item(ActiveGanttVBWCtl1.MathLib.GetRowIndexByPosition(e.Y)).Key
                    MainGrid.ContextMenu = MainGrid.Resources.Item("mnuRowEdit")
                    MainGrid.ContextMenu.PlacementTarget = MainGrid
                    MainGrid.ContextMenu.IsOpen = True
                End If
            Case E_EVENTTARGET.EVT_SELECTEDTASK, E_EVENTTARGET.EVT_TASK
                If e.Button = E_MOUSEBUTTONS.BTN_RIGHT Then
                    Dim oTask As clsTask
                    e.Cancel = True
                    mp_sEditTaskKey = ActiveGanttVBWCtl1.Tasks.Item(ActiveGanttVBWCtl1.MathLib.GetTaskIndexByPosition(e.X, e.Y)).Key
                    oTask = ActiveGanttVBWCtl1.Tasks.Item(mp_sEditTaskKey)
                    MainGrid.ContextMenu = MainGrid.Resources.Item("mnuTaskEdit")
                    Dim mnuEditTask As MenuItem = MainGrid.ContextMenu.Items(0)
                    Dim mnuConvertToRental As MenuItem = MainGrid.ContextMenu.Items(1)
                    If oTask.Tag = 0 Then
                        mnuConvertToRental.Visibility = Windows.Visibility.Visible
                    Else
                        mnuConvertToRental.Visibility = Windows.Visibility.Collapsed
                    End If
                    If oTask.Tag = 2 Then
                        mnuEditTask.Visibility = Windows.Visibility.Collapsed
                    Else
                        mnuEditTask.Visibility = Windows.Visibility.Visible
                    End If
                End If
        End Select
    End Sub

    Private Sub ActiveGanttVBWCtl1_ControlKeyDown(ByVal sender As System.Object, ByVal e As KeyEventArgs) Handles ActiveGanttVBWCtl1.ControlKeyDown
        If e.KeyCode = Key.F2 Then
            Mode = HPE_ADDMODE.AM_RENTAL
        End If
        If e.KeyCode = Key.F9 Then
            Mode = HPE_ADDMODE.AM_MAINTENANCE
        End If
    End Sub

    Private Sub ActiveGanttVBWCtl1_ControlKeyUp(ByVal sender As System.Object, ByVal e As KeyEventArgs) Handles ActiveGanttVBWCtl1.ControlKeyUp
        Mode = HPE_ADDMODE.AM_RESERVATION
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

#End Region

#Region "Form Properties"

    Public Property Mode() As HPE_ADDMODE
        Get
            Return mp_yAddMode
        End Get
        Set(ByVal Value As HPE_ADDMODE)
            mp_yAddMode = Value
            Select Case mp_yAddMode
                Case HPE_ADDMODE.AM_RESERVATION
                    lblMode.Content = "Add Reservation Mode"
                    lblMode.Foreground = New SolidColorBrush(Color.FromRgb(153, 170, 194))
                    mp_sAddModeStyleIndex = "Reservation"
                Case HPE_ADDMODE.AM_RENTAL
                    lblMode.Content = "Add Rental Mode"
                    lblMode.Foreground = New SolidColorBrush(Color.FromRgb(162, 78, 50))
                    mp_sAddModeStyleIndex = "Rental"
                Case HPE_ADDMODE.AM_MAINTENANCE
                    lblMode.Content = "Add Maintenance Mode"
                    lblMode.Foreground = New SolidColorBrush(Color.FromRgb(255, 77, 1))
                    mp_sAddModeStyleIndex = "Maintenance"
            End Select
        End Set
    End Property

    Private Property Zoom() As Integer
        Get
            Return mp_lZoom
        End Get
        Set(ByVal Value As Integer)
            If Value > 5 Or Value < 1 Then
                Return
            End If
            mp_lZoom = Value
            Dim oView As clsView = Nothing
            oView = ActiveGanttVBWCtl1.CurrentViewObject
            Select Case mp_lZoom
                Case 5
                    oView.Interval = E_INTERVAL.IL_MINUTE
                    oView.Factor = 30
                    oView.ClientArea.Grid.Interval = E_INTERVAL.IL_HOUR
                    oView.ClientArea.Grid.Factor = 12
                    oView.TimeLine.TickMarkArea.Visible = False
                Case 4
                    oView.Interval = E_INTERVAL.IL_MINUTE
                    oView.Factor = 15
                    oView.ClientArea.Grid.Interval = E_INTERVAL.IL_HOUR
                    oView.ClientArea.Grid.Factor = 6
                    oView.TimeLine.TickMarkArea.Visible = False
                Case 3
                    oView.Interval = E_INTERVAL.IL_MINUTE
                    oView.Factor = 10
                    oView.ClientArea.Grid.Interval = E_INTERVAL.IL_HOUR
                    oView.ClientArea.Grid.Factor = 3
                    oView.TimeLine.TickMarkArea.Visible = False
                Case 2
                    oView.Interval = E_INTERVAL.IL_MINUTE
                    oView.Factor = 5
                    oView.ClientArea.Grid.Interval = E_INTERVAL.IL_HOUR
                    oView.ClientArea.Grid.Factor = 2
                    oView.TimeLine.TickMarkArea.Visible = True
                    oView.TimeLine.TickMarkArea.Height = 30
                    oView.TimeLine.TickMarkArea.TickMarks.Clear()
                    oView.TimeLine.TickMarkArea.TickMarks.Add(E_INTERVAL.IL_HOUR, 6, E_TICKMARKTYPES.TLT_BIG, True, "hh:mmtt")
                    oView.TimeLine.TickMarkArea.TickMarks.Add(E_INTERVAL.IL_HOUR, 1, E_TICKMARKTYPES.TLT_SMALL, False, "h")
                Case 1
                    oView.Interval = E_INTERVAL.IL_MINUTE
                    oView.Factor = 1
                    oView.ClientArea.Grid.Interval = E_INTERVAL.IL_MINUTE
                    oView.ClientArea.Grid.Factor = 15
                    oView.TimeLine.TickMarkArea.Visible = True
                    oView.TimeLine.TickMarkArea.Height = 30
                    oView.TimeLine.TickMarkArea.TickMarks.Clear()
                    oView.TimeLine.TickMarkArea.TickMarks.Add(E_INTERVAL.IL_HOUR, 1, E_TICKMARKTYPES.TLT_BIG, True, "hh:mmtt")
            End Select
            ActiveGanttVBWCtl1.Redraw()
        End Set
    End Property

#End Region

#Region "Functions"

    Friend Function GetDescription(ByVal lCarTypeID As Integer) As String
        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            Dim sReturn As String = ""
            Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
                Dim oReader As OleDbDataReader = Nothing
                oReader = g_DST_ACCESS_ReturnReader("SELECT sDescription FROM tb_CR_Car_Types WHERE lCarTypeID = " & lCarTypeID, oConn)
                If oReader.Read = True Then
                    sReturn = oReader.Item("sDescription").ToString()
                End If
                oReader.Close()
            End Using
            Return sReturn
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            Return mp_otb_CR_Car_Types.Tables(1).Rows.Find(lCarTypeID).Item("sDescription").ToString()
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            Return mp_otb_CR_Car_Types.Tables(0).Rows.Find(lCarTypeID).Item("sDescription").ToString()
        End If
        Return ""
    End Function

    Private Sub CalculateRate(ByRef oTask As clsTask)
        Dim fFactor As Single = 0
        Dim sRowTag As String()
        Dim lRate As Single = 0
        Dim fSubTotal As Single = 0
        Dim fOptions As Single = 0
        Dim fSurcharge As Double = 0
        Dim fTax As Double = 0
        Dim cALI As Double = 0
        Dim dCRF As Double = 0
        Dim cERF As Double = 0
        Dim cGPS As Double = 0
        Dim cLDW As Double = 0
        Dim cPAI As Double = 0
        Dim cPEP As Double = 0
        Dim cRCFC As Double = 0
        Dim cVLF As Double = 0
        Dim cWTB As Double = 0
        Dim bGPS As Boolean = False
        Dim bLDW As Boolean = False
        Dim bPAI As Boolean = False
        Dim bPEP As Boolean = False
        Dim bALI As Boolean = False
        Dim sName As String = ""
        Dim sPhone As String = ""

        Dim sEstimatedTotal As String = ""
        Dim cEstimatedTotal As Double = 0

        Dim oDataRow As DataRow = Nothing

        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
                Dim oReader As OleDbDataReader = Nothing
                oReader = g_DST_ACCESS_ReturnReader("SELECT * FROM tb_CR_Rentals WHERE lTaskID = " & oTask.Key.Replace("K", ""), oConn)
                If oReader.Read = True Then
                    sName = DirectCast(oReader.Item("sName"), System.String)
                    sPhone = DirectCast(oReader.Item("sPhone"), System.String)

                    bGPS = DirectCast(oReader.Item("bGPS"), System.Boolean)
                    cGPS = CType(oReader.Item("cGPS"), System.Double)
                    bLDW = DirectCast(oReader.Item("bLDW"), System.Boolean)
                    cLDW = CType(oReader.Item("cLDW"), System.Double)
                    bPAI = DirectCast(oReader.Item("bPAI"), System.Boolean)
                    cPAI = CType(oReader.Item("cPAI"), System.Double)
                    bPEP = DirectCast(oReader.Item("bPEP"), System.Boolean)
                    cPEP = CType(oReader.Item("cPEP"), System.Double)
                    bALI = DirectCast(oReader.Item("bALI"), System.Boolean)
                    cALI = CType(oReader.Item("cALI"), System.Double)

                    cERF = CType(oReader.Item("cERF"), System.Double)
                    cWTB = CType(oReader.Item("cWTB"), System.Double)
                    cRCFC = CType(oReader.Item("cRCFC"), System.Double)
                    cVLF = CType(oReader.Item("cVLF"), System.Double)
                    dCRF = CType(oReader.Item("dCRF"), System.Double)
                End If
                oReader.Close()
            End Using
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then

            oDataRow = mp_otb_CR_Rentals.Tables(1).Rows.Find(oTask.Key.Replace("K", ""))
            sName = DirectCast(oDataRow("sName"), System.String)
            sPhone = DirectCast(oDataRow("sPhone"), System.String)

            bGPS = DirectCast(oDataRow("bGPS"), System.Boolean)
            cGPS = DirectCast(oDataRow("cGPS"), System.Double)
            bLDW = DirectCast(oDataRow("bLDW"), System.Boolean)
            cLDW = DirectCast(oDataRow("cLDW"), System.Double)
            bPAI = DirectCast(oDataRow("bPAI"), System.Boolean)
            cPAI = DirectCast(oDataRow("cPAI"), System.Double)
            bPEP = DirectCast(oDataRow("bPEP"), System.Boolean)
            cPEP = DirectCast(oDataRow("cPEP"), System.Double)
            bALI = DirectCast(oDataRow("bALI"), System.Boolean)
            cALI = DirectCast(oDataRow("cALI"), System.Double)

            cERF = DirectCast(oDataRow("cERF"), System.Double)
            cWTB = DirectCast(oDataRow("cWTB"), System.Double)
            cRCFC = DirectCast(oDataRow("cRCFC"), System.Double)
            cVLF = DirectCast(oDataRow("cVLF"), System.Double)
            dCRF = DirectCast(oDataRow("dCRF"), System.Double)
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            oDataRow = mp_otb_CR_Rentals.Tables(0).Rows.Find(oTask.Key.Replace("K", ""))
            sName = DirectCast(oDataRow("sName"), System.String)
            sPhone = DirectCast(oDataRow("sPhone"), System.String)

            bGPS = DirectCast(oDataRow("bGPS"), System.Boolean)
            cGPS = DirectCast(oDataRow("cGPS"), System.Double)
            bLDW = DirectCast(oDataRow("bLDW"), System.Boolean)
            cLDW = DirectCast(oDataRow("cLDW"), System.Double)
            bPAI = DirectCast(oDataRow("bPAI"), System.Boolean)
            cPAI = DirectCast(oDataRow("cPAI"), System.Double)
            bPEP = DirectCast(oDataRow("bPEP"), System.Boolean)
            cPEP = DirectCast(oDataRow("cPEP"), System.Double)
            bALI = DirectCast(oDataRow("bALI"), System.Boolean)
            cALI = DirectCast(oDataRow("cALI"), System.Double)

            cERF = DirectCast(oDataRow("cERF"), System.Double)
            cWTB = DirectCast(oDataRow("cWTB"), System.Double)
            cRCFC = DirectCast(oDataRow("cRCFC"), System.Double)
            cVLF = DirectCast(oDataRow("cVLF"), System.Double)
            dCRF = DirectCast(oDataRow("dCRF"), System.Double)
        End If

        fFactor = CType(ActiveGanttVBWCtl1.MathLib.DateTimeDiff(E_INTERVAL.IL_HOUR, oTask.StartDate, oTask.EndDate) / 24, System.Single)

        If bGPS = True Then
            cGPS = cGPS * fFactor
        Else
            cGPS = 0
        End If
        If bLDW = True Then
            cLDW = cLDW * fFactor
        Else
            cLDW = 0
        End If
        If bPAI = True Then
            cPAI = cPAI * fFactor
        Else
            cPAI = 0
        End If
        If bPEP = True Then
            cPEP = cPEP * fFactor
        Else
            cPEP = 0
        End If
        If bALI = True Then
            cALI = cALI * fFactor
        Else
            cALI = 0
        End If
        sRowTag = oTask.Row.Tag.Split("|"c)
        lRate = CType(sRowTag(1), System.Single)
        cERF = cERF * fFactor
        cWTB = cWTB * fFactor
        cRCFC = cRCFC * fFactor
        cVLF = cVLF * fFactor
        dCRF = dCRF * lRate * fFactor
        fSurcharge = cERF + cWTB + cRCFC + cVLF + dCRF
        fOptions = CType(cGPS + cLDW + cPAI + cPEP + cALI, System.Single)
        fSubTotal = CType(fSurcharge + (lRate * fFactor), System.Single)
        Dim sState As String = ""
        fTax = fSubTotal * GetStateTax(oTask, sState)
        cEstimatedTotal = fSubTotal + fTax + fOptions
        sEstimatedTotal = cEstimatedTotal.ToString("0.00")
        If oTask.Tag = "0" Or oTask.Tag = "1" Then
            oTask.Text = sName & vbCrLf & "Phone: " & sPhone & vbCrLf & "Estimated Total: " & sEstimatedTotal & " USD"
        Else
            cEstimatedTotal = 0
            lRate = 0
        End If

        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            Dim oDB As New clsDB()
            oDB.AddParameter("dtPickUp", oTask.StartDate, clsDB.ParamType.PT_DATE)
            oDB.AddParameter("dtReturn", oTask.EndDate, clsDB.ParamType.PT_DATE)
            oDB.AddParameter("cRate", lRate, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("cEstimatedTotal", cEstimatedTotal, clsDB.ParamType.PT_NUMERIC)
            oDB.ExecuteUpdate("tb_CR_Rentals", "lTaskID = " & oTask.Key.Replace("K", ""))
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            oDataRow = mp_otb_CR_Rentals.Tables(1).Rows.Find(oTask.Key.Replace("K", ""))
            oDataRow("dtPickUp") = oTask.StartDate.DateTimePart
            oDataRow("dtReturn") = oTask.EndDate.DateTimePart
            oDataRow("cRate") = lRate
            oDataRow("cEstimatedTotal") = cEstimatedTotal
            mp_otb_CR_Rentals.WriteXml(g_GetAppLocation() & "\CR_XML\tb_CR_Rentals.xml")
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            oDataRow = mp_otb_CR_Rentals.Tables(0).Rows.Find(oTask.Key.Replace("K", ""))
            oDataRow("dtPickUp") = oTask.StartDate.DateTimePart
            oDataRow("dtReturn") = oTask.EndDate.DateTimePart
            oDataRow("cRate") = lRate
            oDataRow("cEstimatedTotal") = cEstimatedTotal
        End If
    End Sub

    Friend Function GetStateTax(ByRef oTask As clsTask, ByRef sState As String) As Double
        Dim oNode As clsNode = Nothing
        Dim dTax As Double = 0
        Dim oDataRow As DataRow = Nothing
        oNode = oTask.Row.Node.Parent()
        If oNode Is Nothing Then
            Return 0.1
        Else
            If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
                    Dim oReader As OleDbDataReader = Nothing
                    oReader = g_DST_ACCESS_ReturnReader("SELECT sState FROM tb_CR_Rows WHERE lRowID = " & oNode.Row.Key.Replace("K", ""), oConn)
                    If oReader.Read = True Then
                        sState = oReader.Item("sState").ToString()
                    End If
                    oReader.Close()
                    oReader = g_DST_ACCESS_ReturnReader("SELECT dCarRentalTax FROM tb_CR_US_States WHERE ID = '" & sState & "'", oConn)
                    If oReader.Read = True Then
                        dTax = DirectCast(oReader.Item("dCarRentalTax"), System.Double)
                    End If
                    oReader.Close()
                    oConn.Close()
                End Using
                Return dTax
            ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                oDataRow = mp_otb_CR_Rows.Tables(1).Rows.Find(oNode.Row.Key.Replace("K", ""))
                sState = oDataRow("sState").ToString()
                oDataRow = mp_otb_CR_US_States.Tables(1).Rows.Find(sState)
                dTax = DirectCast(oDataRow("dCarRentalTax"), System.Double)
                Return dTax
            ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                oDataRow = mp_otb_CR_Rows.Tables(0).Rows.Find(oNode.Row.Key.Replace("K", ""))
                sState = oDataRow("sState").ToString()
                oDataRow = mp_otb_CR_US_States.Tables(0).Rows.Find(sState)
                dTax = DirectCast(oDataRow("dCarRentalTax"), System.Double)
                Return dTax
            End If
        End If
    End Function

    Private Function GetImageGIF(ByVal sImage As String) As Image
        Dim oDecoder As New GifBitmapDecoder(GetURI(sImage), BitmapCreateOptions.None, BitmapCacheOption.None)
        Dim oBitmap As BitmapSource = oDecoder.Frames(0)
        Dim oReturn As New Image
        oReturn.Source = oBitmap
        Return oReturn
    End Function

    Private Function GetImage(ByVal sImage As String) As Image
        Dim oDecoder As New JpegBitmapDecoder(GetURI(sImage), BitmapCreateOptions.None, BitmapCacheOption.None)
        Dim oBitmap As BitmapSource = oDecoder.Frames(0)
        Dim oReturn As New Image
        oReturn.Source = oBitmap
        Return oReturn
    End Function

    Private Function GetURI(ByVal sImage As String) As Uri
        Dim oURI As Uri = Nothing
        oURI = New Uri(g_GetAppLocation() & "\" & sImage, UriKind.RelativeOrAbsolute)
        Return oURI
    End Function

#End Region

#Region "Toolbar Buttons"

    Private Sub cmdSaveXML_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdSaveXML.Click
        SaveXML()
    End Sub

    Private Sub cmdLoadXML_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdLoadXML.Click
        LoadXML()
    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdPrint.Click
        Print()
    End Sub

    Private Sub cmdZoomIn_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdZoomIn.Click
        Zoom = Zoom - 1
    End Sub

    Private Sub cmdZoomOut_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdZoomOut.Click
        Zoom = Zoom + 1
    End Sub

    Private Sub cmdAddVehicle_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdAddVehicle.Click
        Dim oForm As New fCarRentalVehicle(PRG_DIALOGMODE.DM_ADD, Me, "")
        oForm.ShowDialog()
    End Sub

    Private Sub cmdAddBranch_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdAddBranch.Click
        Dim oForm As New fCarRentalBranch(PRG_DIALOGMODE.DM_ADD, Me, "")
        oForm.ShowDialog()
    End Sub

    Private Sub cmdBack2_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdBack2.Click
        ActiveGanttVBWCtl1.CurrentViewObject.TimeLine.Position(ActiveGanttVBWCtl1.MathLib.DateTimeAdd(ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Grid.Interval, -10 * ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Grid.Factor, ActiveGanttVBWCtl1.CurrentViewObject.TimeLine.StartDate))
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub cmdBack1_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdBack1.Click
        ActiveGanttVBWCtl1.CurrentViewObject.TimeLine.Position(ActiveGanttVBWCtl1.MathLib.DateTimeAdd(ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Grid.Interval, -5 * ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Grid.Factor, ActiveGanttVBWCtl1.CurrentViewObject.TimeLine.StartDate))
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub cmdBack0_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdBack0.Click
        ActiveGanttVBWCtl1.CurrentViewObject.TimeLine.Position(ActiveGanttVBWCtl1.MathLib.DateTimeAdd(ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Grid.Interval, -1 * ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Grid.Factor, ActiveGanttVBWCtl1.CurrentViewObject.TimeLine.StartDate))
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub cmdFwd0_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdFwd0.Click
        ActiveGanttVBWCtl1.CurrentViewObject.TimeLine.Position(ActiveGanttVBWCtl1.MathLib.DateTimeAdd(ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Grid.Interval, 1 * ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Grid.Factor, ActiveGanttVBWCtl1.CurrentViewObject.TimeLine.StartDate))
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub cmdFwd1_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdFwd1.Click
        ActiveGanttVBWCtl1.CurrentViewObject.TimeLine.Position(ActiveGanttVBWCtl1.MathLib.DateTimeAdd(ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Grid.Interval, 5 * ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Grid.Factor, ActiveGanttVBWCtl1.CurrentViewObject.TimeLine.StartDate))
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub cmdFwd2_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdFwd2.Click
        ActiveGanttVBWCtl1.CurrentViewObject.TimeLine.Position(ActiveGanttVBWCtl1.MathLib.DateTimeAdd(ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Grid.Interval, 10 * ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Grid.Factor, ActiveGanttVBWCtl1.CurrentViewObject.TimeLine.StartDate))
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdHelp.Click
        Me.Cursor = Cursors.Wait
        System.Diagnostics.Process.Start("http://www.sourcecodestore.com/Article.aspx?ID=16")
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

    Private Sub mnuPrint_Click1(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuPrint.Click
        Print()
    End Sub

    Private Sub mnuClose_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles mnuClose.Click
        Me.Close()
    End Sub

    Private Sub mnuEditRow_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        Dim oRow As clsRow
        oRow = ActiveGanttVBWCtl1.Rows.Item(mp_sEditRowKey)
        If oRow.Node.Depth = 1 Then
            Dim oForm As New fCarRentalVehicle(PRG_DIALOGMODE.DM_EDIT, Me, mp_sEditRowKey.Replace("K", ""))
            oForm.ShowDialog()
        ElseIf oRow.Node.Depth = 0 Then
            Dim oForm As New fCarRentalBranch(PRG_DIALOGMODE.DM_EDIT, Me, mp_sEditRowKey.Replace("K", ""))
            oForm.ShowDialog()
        End If
    End Sub

    Private Sub mnuDeleteRow_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        If MsgBox("Are you sure you want to delete this item?", MsgBoxStyle.YesNoCancel, "Delete Row") = MsgBoxResult.Yes Then
            Dim oDataRow As DataRow = Nothing
            Dim oDeleteDataRows() As DataRow

            If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                DST_ACCESS.g_DST_ACCESS_ExecuteNonQuery("DELETE * FROM tb_CR_Rows WHERE lRowID = " & mp_sEditRowKey.Replace("K", ""))
                DST_ACCESS.g_DST_ACCESS_ExecuteNonQuery("DELETE * FROM tb_CR_Rentals WHERE lRowID = " & mp_sEditRowKey.Replace("K", ""))
            ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                oDataRow = mp_otb_CR_Rows.Tables(1).Rows.Find(mp_sEditRowKey.Replace("K", ""))
                mp_otb_CR_Rows.Tables(1).Rows.Remove(oDataRow)
                oDeleteDataRows = mp_otb_CR_Rentals.Tables(1).Select("lRowID =" & mp_sEditRowKey.Replace("K", ""))
                For Each oDataRow In oDeleteDataRows
                    mp_otb_CR_Rentals.Tables(1).Rows.Remove(oDataRow)
                Next
                mp_otb_CR_Rows.WriteXml(g_GetAppLocation() & "\CR_XML\tb_CR_Rows.xml")
                mp_otb_CR_Rentals.WriteXml(g_GetAppLocation() & "\CR_XML\tb_CR_Rentals.xml")
            ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                oDataRow = mp_otb_CR_Rows.Tables(0).Rows.Find(mp_sEditRowKey.Replace("K", ""))
                mp_otb_CR_Rows.Tables(0).Rows.Remove(oDataRow)
                oDeleteDataRows = mp_otb_CR_Rentals.Tables(0).Select("lRowID =" & mp_sEditRowKey.Replace("K", ""))
                For Each oDataRow In oDeleteDataRows
                    mp_otb_CR_Rentals.Tables(0).Rows.Remove(oDataRow)
                Next
            End If
            ActiveGanttVBWCtl1.Rows.Remove(mp_sEditRowKey)
            ActiveGanttVBWCtl1.Rows.UpdateTree()
            ActiveGanttVBWCtl1.Redraw()
        End If
    End Sub

    Private Sub mnuEditTask_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        Dim oForm As New fCarRentalReservation(PRG_DIALOGMODE.DM_EDIT, Me, mp_sEditTaskKey.Replace("K", ""))
        oForm.ShowDialog()
    End Sub

    Private Sub mnuConvertToRental_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        Dim oDataRow As DataRow = Nothing
        Dim oTask As clsTask
        oTask = ActiveGanttVBWCtl1.Tasks.Item(mp_sEditTaskKey)
        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            g_DST_ACCESS_ExecuteNonQuery("UPDATE tb_CR_Rentals SET [yMode] = 1 WHERE lTasKID = " & mp_sEditTaskKey.Replace("K", ""))
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            oDataRow = mp_otb_CR_Rentals.Tables(1).Rows.Find(mp_sEditTaskKey.Replace("K", ""))
            oDataRow("yMode") = 1
            mp_otb_CR_Rentals.WriteXml(g_GetAppLocation() & "\CR_XML\tb_CR_Rentals.xml")
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            oDataRow = mp_otb_CR_Rentals.Tables(0).Rows.Find(mp_sEditTaskKey.Replace("K", ""))
            oDataRow("yMode") = 1
        End If
        oTask.Tag = "1"
        oTask.StyleIndex = "Rental"
        ActiveGanttVBWCtl1.Redraw()
    End Sub

    Private Sub mnuDeleteTask_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        If MsgBox("Are you sure you want to delete this item?", MsgBoxStyle.YesNoCancel, "Delete Task") = MsgBoxResult.Yes Then
            Dim oDataRow As DataRow = Nothing
            If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                g_DST_ACCESS_ExecuteNonQuery("DELETE * FROM tb_CR_Rentals WHERE lTaskID = " & mp_sEditTaskKey.Replace("K", ""))
            ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                oDataRow = mp_otb_CR_Rentals.Tables(1).Rows.Find(mp_sEditTaskKey.Replace("K", ""))
                mp_otb_CR_Rentals.Tables(1).Rows.Remove(oDataRow)
                mp_otb_CR_Rentals.WriteXml(g_GetAppLocation() & "\CR_XML\tb_CR_Rentals.xml")
            ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                oDataRow = mp_otb_CR_Rentals.Tables(0).Rows.Find(mp_sEditTaskKey.Replace("K", ""))
                mp_otb_CR_Rentals.Tables(0).Rows.Remove(oDataRow)
            End If
            ActiveGanttVBWCtl1.Tasks.Remove(mp_sEditTaskKey)
            ActiveGanttVBWCtl1.Redraw()
        End If
    End Sub

#End Region

#Region "Toolbar Button & Menu Item Functions"

    Private Sub SaveXML()
        Dim dlg As New Microsoft.Win32.SaveFileDialog()
        dlg.FileName = "AGVBW_CR"
        dlg.DefaultExt = ".xml"
        dlg.Filter = "XML Files (.xml)|*.xml"
        If dlg.ShowDialog() = True Then
            ActiveGanttVBWCtl1.WriteXML(dlg.FileName)
        End If
    End Sub

    Private Sub Print()
        Dim oForm As New fPrintDialog(ActiveGanttVBWCtl1, New AGVBW.DateTime(2009, 6, 1, 0, 0, 0), New AGVBW.DateTime(2009, 6, 30, 0, 0, 0))
        oForm.ShowDialog()
    End Sub

    Private Sub LoadXML()
        Dim oForm As New fLoadXML()
        oForm.ShowDialog()
    End Sub

#End Region

#Region "Load Data"

    Private Sub Access_LoadRowsAndTasks()
        Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
            Dim oCmd As OleDbCommand = Nothing
            Dim oReader As OleDbDataReader = Nothing
            Dim sRowID As String = ""
            Dim oRow As clsRow = Nothing
            Dim oTask As clsTask = Nothing
            oConn.Open()
            oCmd = New OleDbCommand("SELECT * FROM tb_CR_Rows ORDER BY lOrder", oConn)
            oReader = oCmd.ExecuteReader
            While oReader.Read = True
                sRowID = "K" & oReader.Item("lRowID").ToString()
                oRow = ActiveGanttVBWCtl1.Rows.Add(sRowID)
                If DirectCast(oReader.Item("lDepth"), System.Int32) = 0 Then
                    oRow.Text = oReader.Item("sBranchName").ToString() & ", " & oReader.Item("sState").ToString() & vbCrLf & "Phone: " & oReader.Item("sPhone").ToString()
                    oRow.MergeCells = True
                    oRow.Container = False
                    oRow.StyleIndex = "Branch"
                    oRow.ClientAreaStyleIndex = "BranchCA"
                    oRow.Node.Depth = 0
                    oRow.UseNodeImages = True
                    oRow.Node.ExpandedImage = GetImageGIF("CarRental\minus.gif")
                    oRow.Node.Image = GetImageGIF("CarRental\plus.gif")
                    oRow.AllowMove = False
                    oRow.AllowSize = False
                ElseIf DirectCast(oReader.Item("lDepth"), System.Int32) = 1 Then
                    Dim sDescription As String = ""
                    sDescription = GetDescription(DirectCast(oReader.Item("lCarTypeID"), System.Int32))
                    oRow.Cells.Item("1").Text = oReader.Item("sLicensePlates").ToString()
                    oRow.Cells.Item("2").Image = GetImage("CarRental\Small\" & sDescription & ".jpg")
                    oRow.Cells.Item("3").Text = sDescription & vbCrLf & oReader.Item("sACRISSCode").ToString() & " - " & oReader.Item("cRate").ToString() & " USD"
                    oRow.Node.Depth = 1
                    oRow.Tag = oReader.Item("sACRISSCode").ToString() & "|" & oReader.Item("cRate").ToString() & "|" & oReader.Item("lCarTypeID").ToString()
                End If
            End While
            oReader.Close()
            oCmd = New OleDbCommand("SELECT * FROM tb_CR_Rentals", oConn)
            oReader = oCmd.ExecuteReader()
            While oReader.Read = True
                oTask = ActiveGanttVBWCtl1.Tasks.Add("", "K" & oReader.Item("lRowID").ToString(), FromDate(oReader.Item("dtPickUp")), FromDate(oReader.Item("dtReturn")), "K" & oReader.Item("lTaskID").ToString())
                If DirectCast(oReader.Item("yMode"), System.Int16) = 2 Then
                    oTask.Text = "Scheduled Maintenance"
                    oTask.StyleIndex = "Maintenance"
                Else
                    oTask.Text = oReader.Item("sName").ToString() & vbCrLf & "Phone: " & oReader.Item("sPhone").ToString() & vbCrLf & "Estimated Total: " & g_Format(CType(oReader.Item("cEstimatedTotal"), System.Double), "0.00") & " USD"
                    If DirectCast(oReader.Item("yMode"), System.Int16) = 0 Then
                        oTask.StyleIndex = "Reservation"
                    ElseIf DirectCast(oReader.Item("yMode"), System.Int16) = 1 Then
                        oTask.StyleIndex = "Rental"
                    End If
                End If
                oTask.Tag = oReader.Item("yMode").ToString()
            End While
            oReader.Close()
        End Using
    End Sub

    Private Sub XML_LoadRowsAndTasks()
        Dim sRowID As String = ""
        Dim oRow As clsRow = Nothing
        Dim oTask As clsTask = Nothing
        Dim oKeys_tb_CR_Rows(0) As DataColumn
        Dim oKeys_tb_CR_Rentals(0) As DataColumn

        mp_otb_CR_Rows = New DataSet()
        mp_otb_CR_Rows.ReadXmlSchema(g_GetAppLocation() & "\CR_XML\tb_CR_Rows.xsd")
        mp_otb_CR_Rows.ReadXml(g_GetAppLocation() & "\CR_XML\tb_CR_Rows.xml")
        oKeys_tb_CR_Rows(0) = mp_otb_CR_Rows.Tables(1).Columns("lRowID")
        mp_otb_CR_Rows.Tables(1).PrimaryKey = oKeys_tb_CR_Rows

        For Each oDataRow As DataRow In mp_otb_CR_Rows.Tables(1).Rows
            sRowID = "K" & oDataRow("lRowID").ToString()
            oRow = ActiveGanttVBWCtl1.Rows.Add(sRowID)
            If DirectCast(oDataRow("lDepth"), System.Int32) = 0 Then
                oRow.Text = oDataRow("sBranchName").ToString() & ", " & oDataRow("sState").ToString() & vbCrLf & "Phone: " & oDataRow("sPhone").ToString()
                oRow.MergeCells = True
                oRow.Container = False
                oRow.StyleIndex = "Branch"
                oRow.ClientAreaStyleIndex = "BranchCA"
                oRow.Node.Depth = 0
                oRow.UseNodeImages = True
                oRow.Node.ExpandedImage = GetImageGIF("CarRental\minus.gif")
                oRow.Node.Image = GetImageGIF("CarRental\plus.gif")
                oRow.AllowMove = False
                oRow.AllowSize = False
            ElseIf DirectCast(oDataRow("lDepth"), System.Int32) = 1 Then
                Dim sDescription As String = ""
                sDescription = GetDescription(DirectCast(oDataRow("lCarTypeID"), System.Int32))
                oRow.Cells.Item("1").Text = oDataRow("sLicensePlates").ToString()
                oRow.Cells.Item("2").Image = GetImage("CarRental\Small\" & sDescription & ".jpg")
                oRow.Cells.Item("3").Text = sDescription & vbCrLf & oDataRow("sACRISSCode").ToString() & " - " & oDataRow("cRate").ToString() & " USD"
                oRow.Node.Depth = 1
                oRow.Tag = oDataRow("sACRISSCode").ToString() & "|" & oDataRow("cRate").ToString() & "|" & oDataRow("lCarTypeID").ToString()
            End If
        Next

        mp_otb_CR_Rentals = New DataSet()
        mp_otb_CR_Rentals.ReadXmlSchema(g_GetAppLocation() & "\CR_XML\tb_CR_Rentals.xsd")
        mp_otb_CR_Rentals.ReadXml(g_GetAppLocation() & "\CR_XML\tb_CR_Rentals.xml")
        oKeys_tb_CR_Rentals(0) = mp_otb_CR_Rentals.Tables(1).Columns("lTaskID")
        mp_otb_CR_Rentals.Tables(1).PrimaryKey = oKeys_tb_CR_Rentals

        For Each oDataRow As DataRow In mp_otb_CR_Rentals.Tables(1).Rows
            oTask = ActiveGanttVBWCtl1.Tasks.Add("", "K" & oDataRow("lRowID").ToString(), FromDate(oDataRow("dtPickUp")), FromDate(oDataRow("dtReturn")), "K" & oDataRow("lTaskID").ToString())
            If DirectCast(oDataRow("yMode"), System.Int16) = 2 Then
                oTask.Text = "Scheduled Maintenance"
                oTask.StyleIndex = "Maintenance"
            Else
                oTask.Text = oDataRow("sName").ToString() & vbCrLf & "Phone: " & oDataRow("sPhone").ToString() & vbCrLf & "Estimated Total: " & g_Format(DirectCast(oDataRow("cEstimatedTotal"), System.Double), "0.00") & " USD"
                If DirectCast(oDataRow("yMode"), System.Int16) = 0 Then
                    oTask.StyleIndex = "Reservation"
                ElseIf DirectCast(oDataRow("yMode"), System.Int16) = 1 Then
                    oTask.StyleIndex = "Rental"
                End If
            End If
            oTask.Tag = oDataRow("yMode").ToString()
        Next
    End Sub

    Private Sub XML_Load_Car_Types()
        Dim oKeys_tb_CR_Car_Types(0) As DataColumn

        mp_otb_CR_Car_Types = New DataSet()
        mp_otb_CR_Car_Types.ReadXmlSchema(g_GetAppLocation() & "\CR_XML\tb_CR_Car_Types.xsd")
        mp_otb_CR_Car_Types.ReadXml(g_GetAppLocation() & "\CR_XML\tb_CR_Car_Types.xml")

        oKeys_tb_CR_Car_Types(0) = mp_otb_CR_Car_Types.Tables(1).Columns("lCarTypeID")
        mp_otb_CR_Car_Types.Tables(1).PrimaryKey = oKeys_tb_CR_Car_Types
    End Sub

    Private Sub XML_Load_US_States()
        Dim oKeys_tb_CR_US_States(0) As DataColumn

        mp_otb_CR_US_States = New DataSet()
        mp_otb_CR_US_States.ReadXmlSchema(g_GetAppLocation() & "\CR_XML\tb_CR_US_States.xsd")
        mp_otb_CR_US_States.ReadXml(g_GetAppLocation() & "\CR_XML\tb_CR_US_States.xml")
        oKeys_tb_CR_US_States(0) = mp_otb_CR_US_States.Tables(1).Columns("ID")
        mp_otb_CR_US_States.Tables(1).PrimaryKey = oKeys_tb_CR_US_States

    End Sub

    Private Sub XML_Load_ACRISS_Codes()
        Dim oKeys_tb_CR_ACRISS_Codes(0) As DataColumn

        mp_otb_CR_ACRISS_Codes = New DataSet()
        mp_otb_CR_ACRISS_Codes.ReadXmlSchema(g_GetAppLocation() & "\CR_XML\tb_CR_ACRISS_Codes.xsd")
        mp_otb_CR_ACRISS_Codes.ReadXml(g_GetAppLocation() & "\CR_XML\tb_CR_ACRISS_Codes.xml")
        oKeys_tb_CR_ACRISS_Codes(0) = mp_otb_CR_ACRISS_Codes.Tables(1).Columns("ID")
        mp_otb_CR_ACRISS_Codes.Tables(1).PrimaryKey = oKeys_tb_CR_ACRISS_Codes
    End Sub

    Private Sub XML_Load_Taxes_Surcharges_Options()
        Dim oKeys_tb_CR_Taxes_Surcharges_Options(0) As DataColumn

        mp_otb_CR_Taxes_Surcharges_Options = New DataSet()
        mp_otb_CR_Taxes_Surcharges_Options.ReadXmlSchema(g_GetAppLocation() & "\CR_XML\tb_CR_Taxes_Surcharges_Options.xsd")
        mp_otb_CR_Taxes_Surcharges_Options.ReadXml(g_GetAppLocation() & "\CR_XML\tb_CR_Taxes_Surcharges_Options.xml")
        oKeys_tb_CR_Taxes_Surcharges_Options(0) = mp_otb_CR_Taxes_Surcharges_Options.Tables(1).Columns("sID")
        mp_otb_CR_Taxes_Surcharges_Options.Tables(1).PrimaryKey = oKeys_tb_CR_Taxes_Surcharges_Options
    End Sub

    Private Sub NoDataSource_LoadRowsAndTasks()
        Dim sRowID As String = ""
        Dim oRow As clsRow = Nothing
        Dim oTask As clsTask = Nothing
        Dim oKeys_tb_CR_Rows(0) As DataColumn
        Dim oTable_tb_CR_Rows As New DataTable("tb_CR_Rows")
        Dim oKeys_tb_CR_Rentals(0) As DataColumn
        Dim oTable_tb_CR_Rentals As New DataTable("tb_CR_Rentals")

        oTable_tb_CR_Rows.Columns.Add("lRowID", Type.GetType("System.Int32"))
        oTable_tb_CR_Rows.Columns.Add("lDepth", Type.GetType("System.Int32"))
        oTable_tb_CR_Rows.Columns.Add("lOrder", Type.GetType("System.Int32"))
        oTable_tb_CR_Rows.Columns.Add("sLicensePlates", Type.GetType("System.String"))
        oTable_tb_CR_Rows.Columns.Add("lCarTypeID", Type.GetType("System.Int32"))
        oTable_tb_CR_Rows.Columns.Add("sNotes", Type.GetType("System.String"))
        oTable_tb_CR_Rows.Columns.Add("cRate", Type.GetType("System.Double"))
        oTable_tb_CR_Rows.Columns.Add("sACRISSCode", Type.GetType("System.String"))
        oTable_tb_CR_Rows.Columns.Add("sCity", Type.GetType("System.String"))
        oTable_tb_CR_Rows.Columns.Add("sBranchName", Type.GetType("System.String"))
        oTable_tb_CR_Rows.Columns.Add("sState", Type.GetType("System.String"))
        oTable_tb_CR_Rows.Columns.Add("sPhone", Type.GetType("System.String"))
        oTable_tb_CR_Rows.Columns.Add("sManagerName", Type.GetType("System.String"))
        oTable_tb_CR_Rows.Columns.Add("sManagerMobile", Type.GetType("System.String"))
        oTable_tb_CR_Rows.Columns.Add("sAddress", Type.GetType("System.String"))
        oTable_tb_CR_Rows.Columns.Add("sZIP", Type.GetType("System.String"))

        mp_otb_CR_Rows = New DataSet()
        mp_otb_CR_Rows.Tables().Add(oTable_tb_CR_Rows)

        oKeys_tb_CR_Rows(0) = mp_otb_CR_Rows.Tables(0).Columns("lRowID")
        mp_otb_CR_Rows.Tables(0).PrimaryKey = oKeys_tb_CR_Rows

        NoDataSorce_Load_Rows(oTable_tb_CR_Rows)

        For Each oDataRow As DataRow In mp_otb_CR_Rows.Tables(0).Rows
            sRowID = "K" & oDataRow("lRowID").ToString()
            oRow = ActiveGanttVBWCtl1.Rows.Add(sRowID)
            oRow.AllowTextEdit = True
            If DirectCast(oDataRow("lDepth"), System.Int32) = 0 Then
                oRow.Text = oDataRow("sBranchName").ToString() & ", " & oDataRow("sState").ToString() & vbCrLf & "Phone: " & oDataRow("sPhone").ToString()
                oRow.MergeCells = True
                oRow.Container = False
                oRow.StyleIndex = "Branch"
                oRow.ClientAreaStyleIndex = "BranchCA"
                oRow.Node.Depth = 0
                oRow.UseNodeImages = True
                oRow.Node.ExpandedImage = GetImageGIF("CarRental\minus.gif")
                oRow.Node.Image = GetImageGIF("CarRental\plus.gif")
                oRow.AllowMove = False
                oRow.AllowSize = False
            ElseIf DirectCast(oDataRow("lDepth"), System.Int32) = 1 Then
                Dim sDescription As String = ""
                sDescription = GetDescription(DirectCast(oDataRow("lCarTypeID"), System.Int32))
                oRow.Cells.Item("1").Text = oDataRow("sLicensePlates").ToString()
                oRow.Cells.Item("1").AllowTextEdit = True
                oRow.Cells.Item("2").Image = GetImage("CarRental\Small\" & sDescription & ".jpg")
                oRow.Cells.Item("3").Text = sDescription & vbCrLf & oDataRow("sACRISSCode").ToString() & " - " & oDataRow("cRate").ToString() & " USD"
                oRow.Cells.Item("3").AllowTextEdit = True
                oRow.Node.Depth = 1
                oRow.Tag = oDataRow("sACRISSCode").ToString() & "|" & oDataRow("cRate").ToString() & "|" & oDataRow("lCarTypeID").ToString()
            End If
        Next

        oTable_tb_CR_Rentals.Columns.Add("lTaskID", Type.GetType("System.Int32"))
        oTable_tb_CR_Rentals.Columns.Add("lRowID", Type.GetType("System.Int32"))
        oTable_tb_CR_Rentals.Columns.Add("yMode", Type.GetType("System.Int16"))
        oTable_tb_CR_Rentals.Columns.Add("sName", Type.GetType("System.String"))
        oTable_tb_CR_Rentals.Columns.Add("sAddress", Type.GetType("System.String"))
        oTable_tb_CR_Rentals.Columns.Add("sCity", Type.GetType("System.String"))
        oTable_tb_CR_Rentals.Columns.Add("sState", Type.GetType("System.String"))
        oTable_tb_CR_Rentals.Columns.Add("sZIP", Type.GetType("System.String"))
        oTable_tb_CR_Rentals.Columns.Add("sPhone", Type.GetType("System.String"))
        oTable_tb_CR_Rentals.Columns.Add("sMobile", Type.GetType("System.String"))
        oTable_tb_CR_Rentals.Columns.Add("dtPickUp", Type.GetType("System.DateTime"))
        oTable_tb_CR_Rentals.Columns.Add("dtReturn", Type.GetType("System.DateTime"))
        oTable_tb_CR_Rentals.Columns.Add("cRate", Type.GetType("System.Double"))
        oTable_tb_CR_Rentals.Columns.Add("cALI", Type.GetType("System.Double"))
        oTable_tb_CR_Rentals.Columns.Add("dCRF", Type.GetType("System.Double"))
        oTable_tb_CR_Rentals.Columns.Add("cERF", Type.GetType("System.Double"))
        oTable_tb_CR_Rentals.Columns.Add("cGPS", Type.GetType("System.Double"))
        oTable_tb_CR_Rentals.Columns.Add("cLDW", Type.GetType("System.Double"))
        oTable_tb_CR_Rentals.Columns.Add("cPAI", Type.GetType("System.Double"))
        oTable_tb_CR_Rentals.Columns.Add("cPEP", Type.GetType("System.Double"))
        oTable_tb_CR_Rentals.Columns.Add("cRCFC", Type.GetType("System.Double"))
        oTable_tb_CR_Rentals.Columns.Add("cVLF", Type.GetType("System.Double"))
        oTable_tb_CR_Rentals.Columns.Add("cWTB", Type.GetType("System.Double"))
        oTable_tb_CR_Rentals.Columns.Add("dTax", Type.GetType("System.Double"))
        oTable_tb_CR_Rentals.Columns.Add("cEstimatedTotal", Type.GetType("System.Double"))
        oTable_tb_CR_Rentals.Columns.Add("bGPS", Type.GetType("System.Boolean"))
        oTable_tb_CR_Rentals.Columns.Add("bFSO", Type.GetType("System.Boolean"))
        oTable_tb_CR_Rentals.Columns.Add("bLDW", Type.GetType("System.Boolean"))
        oTable_tb_CR_Rentals.Columns.Add("bPAI", Type.GetType("System.Boolean"))
        oTable_tb_CR_Rentals.Columns.Add("bPEP", Type.GetType("System.Boolean"))
        oTable_tb_CR_Rentals.Columns.Add("bALI", Type.GetType("System.Boolean"))

        mp_otb_CR_Rentals = New DataSet()
        mp_otb_CR_Rentals.Tables().Add(oTable_tb_CR_Rentals)

        oKeys_tb_CR_Rentals(0) = mp_otb_CR_Rentals.Tables(0).Columns("lTaskID")
        mp_otb_CR_Rentals.Tables(0).PrimaryKey = oKeys_tb_CR_Rentals

        NoDataSorce_Load_Rentals(oTable_tb_CR_Rentals)

        For Each oDataRow As DataRow In mp_otb_CR_Rentals.Tables(0).Rows
            oTask = ActiveGanttVBWCtl1.Tasks.Add("", "K" & oDataRow("lRowID").ToString(), FromDate(oDataRow("dtPickUp")), FromDate(oDataRow("dtReturn")), "K" & oDataRow("lTaskID").ToString())
            oTask.AllowTextEdit = True
            If DirectCast(oDataRow("yMode"), System.Int16) = 2 Then
                oTask.Text = "Scheduled Maintenance"
                oTask.StyleIndex = "Maintenance"
            Else
                oTask.Text = oDataRow("sName").ToString() & vbCrLf & "Phone: " & oDataRow("sPhone").ToString() & vbCrLf & "Estimated Total: " & g_Format(DirectCast(oDataRow("cEstimatedTotal"), System.Double), "0.00") & " USD"
                If DirectCast(oDataRow("yMode"), System.Int16) = 0 Then
                    oTask.StyleIndex = "Reservation"
                ElseIf DirectCast(oDataRow("yMode"), System.Int16) = 1 Then
                    oTask.StyleIndex = "Rental"
                End If
            End If
            oTask.Tag = oDataRow("yMode").ToString()
        Next

    End Sub

    Private Sub NoDataSorce_Load_Rows(ByVal oDataTable As DataTable)
        NoDataSorce_Load_Row(oDataTable, 28, 0, 1, "", 0, "", 0.0, "", "Hillsboro Beach", "Hillsboro Beach", "FL", "(175) 157-9697", "Nancy Mcatee", "(175) 554-7615", "113 Bueno Drive", "22454")
        NoDataSorce_Load_Row(oDataTable, 29, 1, 2, "CKT-2542", 39, "", 245.0, "FFBV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 30, 1, 3, "XXW-9757", 14, "", 37.0, "EDAZ", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 31, 1, 4, "HGO-6751", 16, "", 37.0, "EDAZ", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 32, 1, 5, "QIZ-1491", 17, "", 37.0, "ECAZ", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 33, 1, 6, "WGN-3159", 46, "", 77.0, "LCAR", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 34, 1, 8, "TJS-5515", 37, "", 245.0, "FFBV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 35, 1, 9, "FPN-9487", 31, "", 37.0, "CDMV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 36, 1, 10, "ENU-2926", 26, "", 45.0, "FWAV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 37, 1, 11, "MND-5686", 11, "", 39.0, "IDAV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 38, 1, 12, "ZZY-1567", 18, "", 37.0, "ECAZ", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 39, 0, 13, "", 0, "", 0.0, "", "Woodville", "Woodville", "OK", "(145) 548-2974", "Matthew Risner", "(145) 679-8583", "8 Navarro Junction", "61614")
        NoDataSorce_Load_Row(oDataTable, 40, 1, 14, "SGL-3748", 24, "", 37.0, "CDAV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 41, 1, 15, "VYW-1478", 43, "", 51.0, "FVAV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 42, 1, 16, "LXV-4412", 27, "", 45.0, "FWAV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 44, 1, 7, "IMU-3364", 23, "", 37.0, "CDAV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 45, 1, 17, "FRG-8842", 30, "", 37.0, "CDMV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 46, 1, 18, "OJQ-8553", 14, "", 37.0, "EDAZ", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 47, 1, 19, "INT-3737", 5, "", 223.0, "PWDV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 48, 1, 20, "USM-8758", 47, "", 77.0, "LCAR", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 49, 1, 21, "RRL-2724", 32, "", 37.0, "CDMV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 50, 1, 22, "EMF-3865", 20, "", 37.0, "CDAV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 51, 1, 23, "SRC-5911", 32, "", 37.0, "CDMV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(oDataTable, 52, 1, 24, "VTN-9768", 3, "", 71.0, "IFBV", "", "", "", "", "", "", "", "")
    End Sub

    Private Sub NoDataSorce_Load_Row(ByVal oDataTable As DataTable, ByVal lRowID As Integer, ByVal lDepth As Integer, ByVal lOrder As Integer, ByVal sLicensePlates As String, ByVal lCarTypeID As Integer, ByVal sNotes As String, ByVal cRate As Double, ByVal sACRISSCode As String, ByVal sCity As String, ByVal sBranchName As String, ByVal sState As String, ByVal sPhone As String, ByVal sManagerName As String, ByVal sManagerMobile As String, ByVal sAddress As String, ByVal sZIP As String)
        Dim oDataRow As DataRow = Nothing
        oDataRow = mp_otb_CR_Rows.Tables(0).NewRow()
        oDataRow("lRowID") = lRowID
        oDataRow("lDepth") = lDepth
        oDataRow("lOrder") = lOrder
        oDataRow("sLicensePlates") = sLicensePlates
        oDataRow("lCarTypeID") = lCarTypeID
        oDataRow("sNotes") = sNotes
        oDataRow("cRate") = cRate
        oDataRow("sACRISSCode") = sACRISSCode
        oDataRow("sCity") = sCity
        oDataRow("sBranchName") = sBranchName
        oDataRow("sState") = sState
        oDataRow("sPhone") = sPhone
        oDataRow("sManagerName") = sManagerName
        oDataRow("sManagerMobile") = sManagerMobile
        oDataRow("sAddress") = sAddress
        oDataRow("sZIP") = sZIP
        mp_otb_CR_Rows.Tables(0).Rows.Add(oDataRow)
    End Sub

    Private Sub NoDataSorce_Load_Rentals(ByVal oDataTable As DataTable)
        NoDataSorce_Load_Rental(oDataTable, 21, 30, 0, "Jeromy Lapham", "33 Mckinley Plaza", "Munds Park", "AZ", "37167", "(532) 463-3173", "(532) 793-8291", New AGVBW.DateTime(2009, 6, 13, 0, 0, 0), New AGVBW.DateTime(2009, 6, 20, 0, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 359.22, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 22, 34, 0, "Colleen Nagle", "21 Graziano Street", "George", "SC", "99234", "(266) 819-5725", "(266) 876-2444", New AGVBW.DateTime(2009, 6, 12, 0, 0, 0), New AGVBW.DateTime(2009, 6, 18, 12, 0, 0), 245.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 1923.27, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 23, 36, 0, "Luisa Farrior", "86 Wiegand Courts", "Dayton", "VA", "79821", "(417) 727-8137", "(417) 974-9449", New AGVBW.DateTime(2009, 6, 10, 12, 0, 0), New AGVBW.DateTime(2009, 6, 26, 0, 0, 0), 45.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 941.21, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 25, 32, 0, "Nancy Sandusky", "4 Babcock Street", "Arlington Heights village", "IL", "37895", "(446) 926-4519", "(446) 552-5686", New AGVBW.DateTime(2009, 6, 9, 12, 0, 0), New AGVBW.DateTime(2009, 6, 18, 12, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 461.85, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 26, 29, 0, "Shawn Kidder", "7 Hynes Street", "Vernon Center", "MN", "71625", "(675) 132-8559", "(675) 568-8572", New AGVBW.DateTime(2009, 6, 19, 0, 0, 0), New AGVBW.DateTime(2009, 6, 25, 0, 0, 0), 245.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 1847.03, True, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 27, 33, 1, "Josephina Kuo", "7 Gruber Stravenue", "North Adams", "MA", "29555", "(585) 968-9925", "(585) 789-1551", New AGVBW.DateTime(2009, 6, 11, 12, 0, 0), New AGVBW.DateTime(2009, 6, 22, 12, 0, 0), 77.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 1081.84, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 28, 35, 1, "Sherie Gebhard", "241 Booth Lock", "Bauxite", "AR", "73573", "(893) 882-9983", "(893) 854-1831", New AGVBW.DateTime(2009, 6, 11, 0, 0, 0), New AGVBW.DateTime(2009, 6, 21, 0, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 513.16, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 29, 29, 0, "Linda Roscoe", "17 Rosenberry Underpass", "Siler", "NC", "23686", "(929) 872-1524", "(929) 546-9944", New AGVBW.DateTime(2009, 6, 11, 0, 0, 0), New AGVBW.DateTime(2009, 6, 17, 0, 0, 0), 245.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 1775.33, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 30, 37, 0, "Matthew Alfred", "298 Burcham Street", "Kivalina", "AK", "88648", "(896) 563-7588", "(896) 973-8419", New AGVBW.DateTime(2009, 6, 11, 12, 0, 0), New AGVBW.DateTime(2009, 6, 20, 12, 0, 0), 39.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 483.01, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 31, 31, 1, "Betty Ballew", "6 Gillespie Drive", "Souris", "ND", "99572", "(718) 942-2143", "(718) 726-7799", New AGVBW.DateTime(2009, 6, 13, 0, 0, 0), New AGVBW.DateTime(2009, 6, 23, 12, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 538.82, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 32, 40, 0, "Jame Josephson", "52 Danford Circle", "Arkport village", "NY", "16792", "(289) 991-7674", "(289) 669-9184", New AGVBW.DateTime(2009, 6, 10, 12, 0, 0), New AGVBW.DateTime(2009, 6, 21, 0, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 538.82, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 33, 42, 0, "Penny Holsinger", "1 Mariano Fields", "Seneca Knolls", "NY", "58312", "(372) 274-7459", "(372) 576-9947", New AGVBW.DateTime(2009, 6, 9, 12, 0, 0), New AGVBW.DateTime(2009, 6, 17, 0, 0, 0), 45.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 455.42, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 34, 44, 0, "Linda Gabaldon", "4 Lewellen Boulevard", "Cypress Lake", "FL", "71862", "(626) 786-3444", "(626) 591-2811", New AGVBW.DateTime(2009, 6, 10, 0, 0, 0), New AGVBW.DateTime(2009, 6, 25, 12, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 795.41, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 35, 46, 0, "Gale Cottingham", "717 Seaton Way", "Worthington borough", "PA", "91136", "(799) 683-3813", "(799) 827-3616", New AGVBW.DateTime(2009, 6, 11, 0, 0, 0), New AGVBW.DateTime(2009, 6, 23, 12, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 641.46, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 36, 42, 0, "Brian Grayson", "4 Eckert Drive", "Dunlap village", "IL", "29184", "(598) 441-2575", "(598) 191-9179", New AGVBW.DateTime(2009, 6, 19, 0, 0, 0), New AGVBW.DateTime(2009, 6, 25, 12, 0, 0), 45.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 394.7, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 38, 45, 1, "Dessie Hoffer", "6 Clay Way", "Monett", "MO", "54761", "(648) 657-9664", "(648) 481-3828", New AGVBW.DateTime(2009, 6, 10, 0, 0, 0), New AGVBW.DateTime(2009, 6, 22, 0, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 615.8, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 39, 41, 1, "Vickie Cartier", "43 Jordan Way", "Williamston", "MI", "92739", "(682) 266-8395", "(682) 745-8184", New AGVBW.DateTime(2009, 6, 9, 12, 0, 0), New AGVBW.DateTime(2009, 6, 19, 12, 0, 0), 51.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 677.78, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 41, 47, 0, "Brian Lenoir", "1 Betts Ridges", "Morrisville", "NC", "11594", "(319) 241-1851", "(319) 571-6978", New AGVBW.DateTime(2009, 6, 11, 12, 0, 0), New AGVBW.DateTime(2009, 6, 19, 12, 0, 0), 223.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 2160.16, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 43, 38, 2, "", "", "", "", "", "", "", New AGVBW.DateTime(2009, 6, 11, 0, 0, 0), New AGVBW.DateTime(2009, 6, 24, 12, 0, 0), 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 44, 49, 0, "Allison Peck", "169 Massa Street", "Waldorf", "MD", "91846", "(679) 847-1487", "(679) 513-3341", New AGVBW.DateTime(2009, 6, 10, 0, 0, 0), New AGVBW.DateTime(2009, 6, 17, 12, 0, 0), 77.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 940.05, False, True, True, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 45, 48, 0, "Tiffany Arce", "6 Spires Street", "Hartford village", "IL", "36615", "(362) 357-2429", "(362) 488-4141", New AGVBW.DateTime(2009, 6, 19, 0, 0, 0), New AGVBW.DateTime(2009, 6, 24, 0, 0, 0), 77.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 491.75, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 46, 51, 0, "Felipe Vantassel", "56 Ormsby Street", "Cheswold", "DE", "49225", "(714) 757-2167", "(714) 378-9745", New AGVBW.DateTime(2009, 6, 14, 0, 0, 0), New AGVBW.DateTime(2009, 6, 24, 12, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 538.82, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(oDataTable, 47, 50, 0, "Patricia Cook", "22 Goulet Drive", "Wesleyville borough", "PA", "15945", "(421) 352-2962", "(421) 682-7189", New AGVBW.DateTime(2009, 6, 19, 0, 0, 0), New AGVBW.DateTime(2009, 6, 25, 12, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 333.56, False, False, False, False, False, False)
    End Sub

    Private Sub NoDataSorce_Load_Rental(ByVal oDataTable As DataTable, ByVal lTaskID As Integer, ByVal lRowID As Integer, ByVal yMode As Integer, ByVal sName As String, ByVal sAddress As String, ByVal sCity As String, ByVal sState As String, ByVal sZIP As String, ByVal sPhone As String, ByVal sMobile As String, ByVal dtPickUp As AGVBW.DateTime, ByVal dtReturn As AGVBW.DateTime, ByVal cRate As Double, ByVal cALI As Double, ByVal dCRF As Double, ByVal cERF As Double, ByVal cGPS As Double, ByVal cLDW As Double, ByVal cPAI As Double, ByVal cPEP As Double, ByVal cRCFC As Double, ByVal cVLF As Double, ByVal cWTB As Double, ByVal dTax As Double, ByVal cEstimatedTotal As Double, ByVal bGPS As Boolean, ByVal bFSO As Boolean, ByVal bLDW As Boolean, ByVal bPAI As Boolean, ByVal bPEP As Boolean, ByVal bALI As Boolean)
        Dim oDataRow As DataRow = Nothing
        oDataRow = mp_otb_CR_Rentals.Tables(0).NewRow()
        oDataRow("lTaskID") = lTaskID
        oDataRow("lRowID") = lRowID
        oDataRow("yMode") = yMode
        oDataRow("sName") = sName
        oDataRow("sAddress") = sAddress
        oDataRow("sCity") = sCity
        oDataRow("sState") = sState
        oDataRow("sZIP") = sZIP
        oDataRow("sPhone") = sPhone
        oDataRow("sMobile") = sMobile
        oDataRow("dtPickUp") = dtPickUp.DateTimePart
        oDataRow("dtReturn") = dtReturn.DateTimePart
        oDataRow("cRate") = cRate
        oDataRow("cALI") = cALI
        oDataRow("dCRF") = dCRF
        oDataRow("cERF") = cERF
        oDataRow("cGPS") = cGPS
        oDataRow("cLDW") = cLDW
        oDataRow("cPAI") = cPAI
        oDataRow("cPEP") = cPEP
        oDataRow("cRCFC") = cRCFC
        oDataRow("cVLF") = cVLF
        oDataRow("cWTB") = cWTB
        oDataRow("dTax") = dTax
        oDataRow("cEstimatedTotal") = cEstimatedTotal
        oDataRow("bGPS") = bGPS
        oDataRow("bFSO") = bFSO
        oDataRow("bLDW") = bLDW
        oDataRow("bPAI") = bPAI
        oDataRow("bPEP") = bPEP
        oDataRow("bALI") = bALI
        mp_otb_CR_Rentals.Tables(0).Rows.Add(oDataRow)
    End Sub

    Private Sub NoDataSource_Load_Car_Types()
        Dim oKeys_tb_CR_Car_Types(0) As DataColumn
        Dim oTable_tb_CR_Car_Types As New DataTable("tb_CR_Car_Types")

        oTable_tb_CR_Car_Types.Columns.Add("lCarTypeID", Type.GetType("System.Int32"))
        oTable_tb_CR_Car_Types.Columns.Add("sDescription", Type.GetType("System.String"))
        oTable_tb_CR_Car_Types.Columns.Add("sACRISSCode", Type.GetType("System.String"))
        oTable_tb_CR_Car_Types.Columns.Add("cStdRate", Type.GetType("System.Double"))

        mp_otb_CR_Car_Types = New DataSet()
        mp_otb_CR_Car_Types.Tables().Add(oTable_tb_CR_Car_Types)

        oKeys_tb_CR_Car_Types(0) = mp_otb_CR_Car_Types.Tables(0).Columns("lCarTypeID")
        mp_otb_CR_Car_Types.Tables(0).PrimaryKey = oKeys_tb_CR_Car_Types

        NoDataSource_Add_Car_Type(1, "Escape Panther Black", "IFBV", 71.0)
        NoDataSource_Add_Car_Type(2, "Escape Hot Red", "IFBV", 71.0)
        NoDataSource_Add_Car_Type(3, "Escape Atlantis Blue", "IFBV", 71.0)
        NoDataSource_Add_Car_Type(4, "Escape Metalic Sand", "IFBV", 71.0)
        NoDataSource_Add_Car_Type(5, "Territory TX RWD Ego", "PWDV", 223.0)
        NoDataSource_Add_Car_Type(6, "Territory TX RWD Kashmir", "PWDV", 223.0)
        NoDataSource_Add_Car_Type(7, "Territory TX RWD Steel", "PWDV", 223.0)
        NoDataSource_Add_Car_Type(8, "Territory TX RWD Silhouette", "PWDV", 223.0)
        NoDataSource_Add_Car_Type(9, "Territory TX RWD Winter White", "PWDV", 223.0)
        NoDataSource_Add_Car_Type(10, "Mondeo LX Sea Grey", "IDAV", 39.0)
        NoDataSource_Add_Car_Type(11, "Mondeo LX Ink Blue", "IDAV", 39.0)
        NoDataSource_Add_Car_Type(12, "Mondeo LX Colorado Red", "IDAV", 39.0)
        NoDataSource_Add_Car_Type(13, "Fiesta CL 5 Door Squeeze", "EDAZ", 37.0)
        NoDataSource_Add_Car_Type(14, "Fiesta CL 5 Door Hydro", "EDAZ", 37.0)
        NoDataSource_Add_Car_Type(15, "Fiesta CL 5 Door Panther Black", "EDAZ", 37.0)
        NoDataSource_Add_Car_Type(16, "Fiesta CL 5 Door Frozen White", "EDAZ", 37.0)
        NoDataSource_Add_Car_Type(17, "Fiesta CL 3 Door Ocean", "ECAZ", 37.0)
        NoDataSource_Add_Car_Type(18, "Fiesta CL 3 Door Hydro", "ECAZ", 37.0)
        NoDataSource_Add_Car_Type(19, "Fiesta CL 3 Door Panther Black", "ECAZ", 37.0)
        NoDataSource_Add_Car_Type(20, "Focus CL Sedan Satin White", "CDAV", 37.0)
        NoDataSource_Add_Car_Type(21, "Focus CL Sedan Titanium Grey", "CDAV", 37.0)
        NoDataSource_Add_Car_Type(22, "Focus CL Sedan Black Sapphire", "CDAV", 37.0)
        NoDataSource_Add_Car_Type(23, "Focus CL Sedan Tango", "CDAV", 37.0)
        NoDataSource_Add_Car_Type(24, "Focus CL Sedan Ocean", "CDAV", 37.0)
        NoDataSource_Add_Car_Type(25, "Falcon XT Wagon Lightning Strike", "FWAV", 45.0)
        NoDataSource_Add_Car_Type(26, "Falcon XT Wagon Silhoutte", "FWAV", 45.0)
        NoDataSource_Add_Car_Type(27, "Falcon XT Wagon Sensation", "FWAV", 45.0)
        NoDataSource_Add_Car_Type(28, "Falcon XT Wagon Vixen", "FWAV", 45.0)
        NoDataSource_Add_Car_Type(29, "Falcon XT Wagon Steel", "FWAV", 45.0)
        NoDataSource_Add_Car_Type(30, "Focus CL Hatch Ocean", "CDMV", 37.0)
        NoDataSource_Add_Car_Type(31, "Focus CL Hatch Black Sapphire", "CDMV", 37.0)
        NoDataSource_Add_Car_Type(32, "Focus CL Hatch Tonic", "CDMV", 37.0)
        NoDataSource_Add_Car_Type(33, "Focus CL Hatch Colorado Red", "CDMV", 37.0)
        NoDataSource_Add_Car_Type(34, "Range Rover HSE Alaska White", "FFBV", 245.0)
        NoDataSource_Add_Car_Type(35, "Range Rover HSE Rimini", "FFBV", 245.0)
        NoDataSource_Add_Car_Type(36, "Range Rover HSE Galway Green", "FFBV", 245.0)
        NoDataSource_Add_Car_Type(37, "Range Rover HSE Buckingham Blue", "FFBV", 245.0)
        NoDataSource_Add_Car_Type(38, "Range Rover HSE Santorini Black", "FFBV", 245.0)
        NoDataSource_Add_Car_Type(39, "Range Rover HSE Zermatt Silver", "FFBV", 245.0)
        NoDataSource_Add_Car_Type(40, "LR3 Rimini Red", "FFBV", 232.0)
        NoDataSource_Add_Car_Type(41, "LR3 Santorini Black", "FFBV", 232.0)
        NoDataSource_Add_Car_Type(42, "LR3 Alaska White", "FFBV", 232.0)
        NoDataSource_Add_Car_Type(43, "Town and Country Modern Blue", "FVAV", 51.0)
        NoDataSource_Add_Car_Type(44, "Town and Country Melbourne Green", "FVAV", 51.0)
        NoDataSource_Add_Car_Type(45, "Town and Country Inferno Red", "FVAV", 51.0)
        NoDataSource_Add_Car_Type(46, "Chrysler 300 Clearwater Blue", "LCAR", 77.0)
        NoDataSource_Add_Car_Type(47, "Chrysler 300 Brilliant Black", "LCAR", 77.0)
        NoDataSource_Add_Car_Type(48, "Chrysler 300 Bright Silver", "LCAR", 77.0)
    End Sub

    Private Sub NoDataSource_Add_Car_Type(ByVal lCarTypeID As Integer, ByVal sDescription As String, ByVal sACRISSCode As String, ByVal cStdRate As Double)
        Dim oDataRow As DataRow = Nothing
        oDataRow = mp_otb_CR_Car_Types.Tables(0).NewRow()
        oDataRow("lCarTypeID") = lCarTypeID
        oDataRow("sDescription") = sDescription
        oDataRow("sACRISSCode") = sACRISSCode
        oDataRow("cStdRate") = cStdRate
        mp_otb_CR_Car_Types.Tables(0).Rows.Add(oDataRow)
    End Sub

    Private Sub NoDataSource_Load_US_States()
        Dim oKeys_tb_CR_US_States(0) As DataColumn
        Dim oTable_tb_CR_US_States As New DataTable("tb_CR_US_States")

        oTable_tb_CR_US_States.Columns.Add("ID", Type.GetType("System.String"))
        oTable_tb_CR_US_States.Columns.Add("Name", Type.GetType("System.String"))
        oTable_tb_CR_US_States.Columns.Add("dCarRentalTax", Type.GetType("System.Double"))


        mp_otb_CR_US_States = New DataSet()
        mp_otb_CR_US_States.Tables().Add(oTable_tb_CR_US_States)

        oKeys_tb_CR_US_States(0) = mp_otb_CR_US_States.Tables(0).Columns("ID")
        mp_otb_CR_US_States.Tables(0).PrimaryKey = oKeys_tb_CR_US_States

        NoDataSource_Add_US_State("AK", "Alaska", 0.0)
        NoDataSource_Add_US_State("AL", "Alabama", 0.01)
        NoDataSource_Add_US_State("AR", "Arkansas", 0.01)
        NoDataSource_Add_US_State("AZ", "Arizona", 0.05)
        NoDataSource_Add_US_State("CA", "California", 0.06)
        NoDataSource_Add_US_State("CO", "Colorado", 0.03)
        NoDataSource_Add_US_State("CT", "Connecticut", 0.06)
        NoDataSource_Add_US_State("DE", "Delaware", 0.02)
        NoDataSource_Add_US_State("FL", "Florida", 0.07)
        NoDataSource_Add_US_State("GA", "Georgia", 0.04)
        NoDataSource_Add_US_State("HI", "Hawaii", 0.05)
        NoDataSource_Add_US_State("IA", "Iowa", 0.04)
        NoDataSource_Add_US_State("ID", "Idaho", 0.03)
        NoDataSource_Add_US_State("IL", "Illinois", 0.02)
        NoDataSource_Add_US_State("IN", "Indiana", 0.08)
        NoDataSource_Add_US_State("KS", "Kansas", 0.06)
        NoDataSource_Add_US_State("KY", "Kentucky", 0.05)
        NoDataSource_Add_US_State("LA", "Louisiana", 0.03)
        NoDataSource_Add_US_State("MA", "Massachusetts", 0.06)
        NoDataSource_Add_US_State("MD", "Maryland", 0.02)
        NoDataSource_Add_US_State("ME", "Maine", 0.03)
        NoDataSource_Add_US_State("MI", "Michigan", 0.04)
        NoDataSource_Add_US_State("MN", "Minnesota", 0.04)
        NoDataSource_Add_US_State("MO", "Missouri", 0.02)
        NoDataSource_Add_US_State("MS", "Mississippi", 0.01)
        NoDataSource_Add_US_State("MT", "Montana", 0.03)
        NoDataSource_Add_US_State("NC", "North Carolina", 0.04)
        NoDataSource_Add_US_State("ND", "North Dakota", 0.08)
        NoDataSource_Add_US_State("NE", "Nebraska", 0.06)
        NoDataSource_Add_US_State("NH", "New Hampshire", 0.07)
        NoDataSource_Add_US_State("NJ", "New Jersey", 0.06)
        NoDataSource_Add_US_State("NM", "New Mexico", 0.03)
        NoDataSource_Add_US_State("NV", "Nevada", 0.02)
        NoDataSource_Add_US_State("NY", "New York", 0.03)
        NoDataSource_Add_US_State("OH", "Ohio", 0.02)
        NoDataSource_Add_US_State("OK", "Oklahoma", 0.03)
        NoDataSource_Add_US_State("OR", "Oregon", 0.04)
        NoDataSource_Add_US_State("PA", "Pennsylvania", 0.05)
        NoDataSource_Add_US_State("RI", "Rhode Island", 0.06)
        NoDataSource_Add_US_State("SC", "South Carolina", 0.05)
        NoDataSource_Add_US_State("SD", "South Dakota", 0.04)
        NoDataSource_Add_US_State("TN", "Tennessee", 0.03)
        NoDataSource_Add_US_State("TX", "Texas", 0.02)
        NoDataSource_Add_US_State("UT", "Utah", 0.05)
        NoDataSource_Add_US_State("VA", "Virginia", 0.06)
        NoDataSource_Add_US_State("VT", "Vermont", 0.05)
        NoDataSource_Add_US_State("WA", "Washington", 0.04)
        NoDataSource_Add_US_State("WI", "Wisconsin", 0.06)
        NoDataSource_Add_US_State("WV", "West Virginia", 0.07)
        NoDataSource_Add_US_State("WY", "Wyoming", 0.08)
    End Sub

    Private Sub NoDataSource_Add_US_State(ByVal ID As String, ByVal Name As String, ByVal dCarRentalTax As Single)
        Dim oDataRow As DataRow = Nothing
        oDataRow = mp_otb_CR_US_States.Tables(0).NewRow()
        oDataRow("ID") = ID
        oDataRow("Name") = Name
        oDataRow("dCarRentalTax") = dCarRentalTax
        mp_otb_CR_US_States.Tables(0).Rows.Add(oDataRow)
    End Sub

    Private Sub NoDataSource_Load_ACRISS_Codes()
        Dim oKeys_tb_CR_ACRISS_Codes(0) As DataColumn
        Dim oTable_tb_CR_ACRISS_Codes As New DataTable("tb_CR_ACRISS_Codes")

        oTable_tb_CR_ACRISS_Codes.Columns.Add("ID", Type.GetType("System.Int32"))
        oTable_tb_CR_ACRISS_Codes.Columns.Add("Letter", Type.GetType("System.String"))
        oTable_tb_CR_ACRISS_Codes.Columns.Add("Description", Type.GetType("System.String"))
        oTable_tb_CR_ACRISS_Codes.Columns.Add("Position", Type.GetType("System.Int32"))

        mp_otb_CR_ACRISS_Codes = New DataSet()
        mp_otb_CR_ACRISS_Codes.Tables().Add(oTable_tb_CR_ACRISS_Codes)

        oKeys_tb_CR_ACRISS_Codes(0) = mp_otb_CR_ACRISS_Codes.Tables(0).Columns("ID")
        mp_otb_CR_ACRISS_Codes.Tables(0).PrimaryKey = oKeys_tb_CR_ACRISS_Codes

        NoDataSource_Add_ACRISS_Code(1, "M", "Mini", 1)
        NoDataSource_Add_ACRISS_Code(2, "N", "Mini Elite", 1)
        NoDataSource_Add_ACRISS_Code(3, "E", "Economy", 1)
        NoDataSource_Add_ACRISS_Code(4, "H", "Economy Elite", 1)
        NoDataSource_Add_ACRISS_Code(5, "C", "Compact", 1)
        NoDataSource_Add_ACRISS_Code(6, "D", "Compact Elite", 1)
        NoDataSource_Add_ACRISS_Code(7, "I", "Intermediate", 1)
        NoDataSource_Add_ACRISS_Code(8, "J", "Intermediate Elite", 1)
        NoDataSource_Add_ACRISS_Code(9, "S", "Standard", 1)
        NoDataSource_Add_ACRISS_Code(10, "R", "Standard Elite", 1)
        NoDataSource_Add_ACRISS_Code(11, "F", "Fullsize", 1)
        NoDataSource_Add_ACRISS_Code(12, "G", "Fullsize Elite", 1)
        NoDataSource_Add_ACRISS_Code(13, "P", "Premium", 1)
        NoDataSource_Add_ACRISS_Code(14, "U", "Premium Elite", 1)
        NoDataSource_Add_ACRISS_Code(15, "L", "Luxury", 1)
        NoDataSource_Add_ACRISS_Code(16, "W", "Luxury Elite", 1)
        NoDataSource_Add_ACRISS_Code(17, "O", "Oversize", 1)
        NoDataSource_Add_ACRISS_Code(18, "X", "Special", 1)
        NoDataSource_Add_ACRISS_Code(19, "B", "2-3 Door", 2)
        NoDataSource_Add_ACRISS_Code(20, "C", "2/4 Door", 2)
        NoDataSource_Add_ACRISS_Code(21, "D", "4-5 Door", 2)
        NoDataSource_Add_ACRISS_Code(22, "W", "Wagon/Estate", 2)
        NoDataSource_Add_ACRISS_Code(23, "V", "Passenger Van", 2)
        NoDataSource_Add_ACRISS_Code(24, "L", "Limousine", 2)
        NoDataSource_Add_ACRISS_Code(25, "S", "Sport", 2)
        NoDataSource_Add_ACRISS_Code(26, "T", "Convertible", 2)
        NoDataSource_Add_ACRISS_Code(27, "F", "SUV", 2)
        NoDataSource_Add_ACRISS_Code(28, "J", "Open Air All Terrain", 2)
        NoDataSource_Add_ACRISS_Code(29, "X", "Special", 2)
        NoDataSource_Add_ACRISS_Code(30, "P", "Pick up Regular Cab", 2)
        NoDataSource_Add_ACRISS_Code(31, "Q", "Pick up Extended Cab", 2)
        NoDataSource_Add_ACRISS_Code(32, "Z", "Special Offer Car", 2)
        NoDataSource_Add_ACRISS_Code(33, "E", "Coupe", 2)
        NoDataSource_Add_ACRISS_Code(34, "M", "Monospace", 2)
        NoDataSource_Add_ACRISS_Code(35, "R", "Recreational Vehicle", 2)
        NoDataSource_Add_ACRISS_Code(36, "H", "Motor Home", 2)
        NoDataSource_Add_ACRISS_Code(37, "Y", "2 Wheel Vehicle", 2)
        NoDataSource_Add_ACRISS_Code(38, "N", "Roadster", 2)
        NoDataSource_Add_ACRISS_Code(39, "G", "Crossover", 2)
        NoDataSource_Add_ACRISS_Code(40, "K", "Commercial Van/Truck", 2)
        NoDataSource_Add_ACRISS_Code(41, "M", "Manual Unspecified Drive", 3)
        NoDataSource_Add_ACRISS_Code(42, "N", "Manual 4WD", 3)
        NoDataSource_Add_ACRISS_Code(43, "C", "Manual AWD", 3)
        NoDataSource_Add_ACRISS_Code(44, "A", "Auto Unspecified Drive", 3)
        NoDataSource_Add_ACRISS_Code(45, "B", "Auto 4WD", 3)
        NoDataSource_Add_ACRISS_Code(46, "D", "Auto AWD", 3)
        NoDataSource_Add_ACRISS_Code(47, "R", "Unspecified Fuel/Power With Air", 4)
        NoDataSource_Add_ACRISS_Code(48, "N", "Unspecified Fuel/Power Without Air", 4)
        NoDataSource_Add_ACRISS_Code(49, "D", "Diesel Air", 4)
        NoDataSource_Add_ACRISS_Code(50, "Q", "Diesel No Air", 4)
        NoDataSource_Add_ACRISS_Code(51, "H", "Hybrid Air", 4)
        NoDataSource_Add_ACRISS_Code(52, "I", "Hybrid No Air", 4)
        NoDataSource_Add_ACRISS_Code(53, "E", "Electric Air", 4)
        NoDataSource_Add_ACRISS_Code(54, "C", "Electric No Air", 4)
        NoDataSource_Add_ACRISS_Code(55, "L", "LPG/Compressed Gas Air", 4)
        NoDataSource_Add_ACRISS_Code(56, "S", "LPG/Compressed Gas No Air", 4)
        NoDataSource_Add_ACRISS_Code(57, "A", "Hydrogen Air", 4)
        NoDataSource_Add_ACRISS_Code(58, "B", "Hydrogen No Air", 4)
        NoDataSource_Add_ACRISS_Code(59, "M", "Multi Fuel/Power Air", 4)
        NoDataSource_Add_ACRISS_Code(60, "F", "Multi Fuel/Power No Air", 4)
        NoDataSource_Add_ACRISS_Code(61, "V", "Petrol Air", 4)
        NoDataSource_Add_ACRISS_Code(62, "Z", "Petrol No Air", 4)
        NoDataSource_Add_ACRISS_Code(63, "U", "Ethanol Air", 4)
        NoDataSource_Add_ACRISS_Code(64, "X", "Ethanol No Air", 4)
    End Sub

    Private Sub NoDataSource_Add_ACRISS_Code(ByVal ID As Integer, ByVal Letter As String, ByVal Description As String, ByVal Position As Integer)
        Dim oDataRow As DataRow = Nothing
        oDataRow = mp_otb_CR_ACRISS_Codes.Tables(0).NewRow()
        oDataRow("ID") = ID
        oDataRow("Letter") = Letter
        oDataRow("Description") = Description
        oDataRow("Position") = Position
        mp_otb_CR_ACRISS_Codes.Tables(0).Rows.Add(oDataRow)
    End Sub

    Private Sub NoDataSource_Load_Taxes_Surcharges_Options()
        Dim oKeys_tb_CR_Taxes_Surcharges_Options(0) As DataColumn
        Dim oTable_tb_CR_Taxes_Surcharges_Options As New DataTable("tb_CR_Taxes_Surcharges_Options")

        oTable_tb_CR_Taxes_Surcharges_Options.Columns.Add("sID", Type.GetType("System.String"))
        oTable_tb_CR_Taxes_Surcharges_Options.Columns.Add("sDescription", Type.GetType("System.String"))
        oTable_tb_CR_Taxes_Surcharges_Options.Columns.Add("sRate", Type.GetType("System.Double"))

        mp_otb_CR_Taxes_Surcharges_Options = New DataSet()
        mp_otb_CR_Taxes_Surcharges_Options.Tables().Add(oTable_tb_CR_Taxes_Surcharges_Options)

        oKeys_tb_CR_Taxes_Surcharges_Options(0) = mp_otb_CR_Taxes_Surcharges_Options.Tables(0).Columns("sID")
        mp_otb_CR_Taxes_Surcharges_Options.Tables(0).PrimaryKey = oKeys_tb_CR_Taxes_Surcharges_Options

        NoDataSource_Add_Taxes_Surcharges_Options("ALI", "Additional Liability Insurance", 14.43)
        NoDataSource_Add_Taxes_Surcharges_Options("CRF", "Concession Recovery Fee", 0.1)
        NoDataSource_Add_Taxes_Surcharges_Options("ERF", "Energy Recovery Fee", 0.67)
        NoDataSource_Add_Taxes_Surcharges_Options("GPS", "GPS", 11.95)
        NoDataSource_Add_Taxes_Surcharges_Options("LDW", "Loss Damage Waiver", 26.99)
        NoDataSource_Add_Taxes_Surcharges_Options("PAI", "Personal Accident Insurance", 4.0)
        NoDataSource_Add_Taxes_Surcharges_Options("PEP", "Personal Effects Protection", 2.95)
        NoDataSource_Add_Taxes_Surcharges_Options("RCFC", "Rental Car Facility Charge", 4.0)
        NoDataSource_Add_Taxes_Surcharges_Options("VLF", "Vehicle License Fee", 0.6)
        NoDataSource_Add_Taxes_Surcharges_Options("WTB", "Waste Tire/Battery", 2.03)
    End Sub

    Private Sub NoDataSource_Add_Taxes_Surcharges_Options(ByVal sID As String, ByVal sDescription As String, ByVal sRate As Double)
        Dim oDataRow As DataRow = Nothing
        oDataRow = mp_otb_CR_Taxes_Surcharges_Options.Tables(0).NewRow()
        oDataRow("sID") = sID
        oDataRow("sDescription") = sDescription
        oDataRow("sRate") = sRate
        mp_otb_CR_Taxes_Surcharges_Options.Tables(0).Rows.Add(oDataRow)
    End Sub

#End Region

End Class
