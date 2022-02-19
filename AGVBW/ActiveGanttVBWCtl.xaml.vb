Imports System.ComponentModel

< _
LicenseProviderAttribute(GetType(RegistryLicenseProvider)), _
System.Runtime.InteropServices.GuidAttribute("B2288103-B606-4A3C-933C-C4C88A9C6898") _
> _
Partial Public Class ActiveGanttVBWCtl
    Inherits UserControl


    Private mp_oLicense As License = Nothing
    <System.ComponentModel.Browsable(False)> Public Rows As clsRows
    <System.ComponentModel.Browsable(False)> Public Tasks As clsTasks
    <System.ComponentModel.Browsable(False)> Public Columns As clsColumns
    <System.ComponentModel.Browsable(False)> Public Styles As clsStyles
    <System.ComponentModel.Browsable(False)> Public Layers As clsLayers
    <System.ComponentModel.Browsable(False)> Public Percentages As clsPercentages
    <System.ComponentModel.Browsable(False)> Public TimeBlocks As clsTimeBlocks
    <System.ComponentModel.Browsable(False)> Public Predecessors As clsPredecessors
    <System.ComponentModel.Browsable(False)> Public Views As clsViews
    <System.ComponentModel.Browsable(False)> Public Splitter As clsSplitter
    <System.ComponentModel.Browsable(False)> Public Printer As clsPrinter
    <System.ComponentModel.Browsable(False)> Public Treeview As clsTreeview
    <System.ComponentModel.Browsable(False)> Public Drawing As clsDrawing
    <System.ComponentModel.Browsable(False)> Public MathLib As clsMath
    <System.ComponentModel.Browsable(False)> Public StrLib As clsString
    <System.ComponentModel.Browsable(False)> Public VerticalScrollBar As clsVerticalScrollBar
    <System.ComponentModel.Browsable(False)> Public HorizontalScrollBar As clsHorizontalScrollBar
    <System.ComponentModel.Browsable(False)> Public TierAppearance As clsTierAppearance
    <System.ComponentModel.Browsable(False)> Public TierFormat As clsTierFormat
    <System.ComponentModel.Browsable(False)> Public ScrollBarSeparator As clsScrollBarSeparator

    Private tmpTimeBlocks As clsTimeBlocks
    Friend MouseKeyboardEvents As clsMouseKeyboardEvents
    Private mp_oCurrentView As clsView
    Friend clsG As clsGraphics
    Private mp_bAllowAdd As Boolean = True
    Private mp_bAllowEdit As Boolean = True
    Private mp_bAllowSplitterMove As Boolean = True
    Private mp_bAllowRowSize As Boolean = True
    Private mp_bAllowRowMove As Boolean = True
    Private mp_bAllowColumnSize As Boolean = True
    Private mp_bAllowColumnMove As Boolean = True
    Private mp_bAllowTimeLineScroll As Boolean = True
    Private mp_bAllowPredecessorAdd As Boolean = True
    Private mp_bDoubleBuffering As Boolean = True
    Private mp_bPropertiesRead As Boolean = False
    Private mp_bEnforcePredecessors As Boolean = False
    Private mp_dtTimeLineStartBuffer As AGVBW.DateTime
    Private mp_dtTimeLineEndBuffer As AGVBW.DateTime
    Private mp_lMinColumnWidth As Integer = 5
    Private mp_lMinRowHeight As Integer = 5
    Private mp_lSelectedTaskIndex As Integer = 0
    Private mp_lSelectedColumnIndex As Integer = 0
    Private mp_lSelectedRowIndex As Integer = 0
    Private mp_lSelectedCellIndex As Integer = 0
    Private mp_lSelectedPercentageIndex As Integer = 0
    Private mp_lSelectedPredecessorIndex As Integer = 0
    Private mp_lTreeviewColumnIndex As Integer = 0
    Private mp_sCurrentLayer As String = "0"
    Private mp_sCurrentView As String = ""
    Private mp_yAddMode As E_ADDMODE = E_ADDMODE.AT_TASKADD
    Private mp_yAddDurationInterval As E_INTERVAL = E_INTERVAL.IL_SECOND
    Private mp_yScrollBarBehaviour As E_SCROLLBEHAVIOUR = E_SCROLLBEHAVIOUR.SB_HIDE
    Private mp_yTimeBlockBehaviour As E_TIMEBLOCKBEHAVIOUR = E_TIMEBLOCKBEHAVIOUR.TBB_ROWEXTENTS
    Private mp_yLayerEnableObjects As E_LAYEROBJECTENABLE = E_LAYEROBJECTENABLE.EC_INCURRENTLAYERONLY
    Private mp_yErrorReports As E_REPORTERRORS = E_REPORTERRORS.RE_MSGBOX
    Private mp_yTierAppearanceScope As E_TIERAPPEARANCESCOPE = E_TIERAPPEARANCESCOPE.TAS_CONTROL
    Private mp_yTierFormatScope As E_TIERFORMATSCOPE = E_TIERFORMATSCOPE.TFS_CONTROL
    Private mp_yPredecessorMode As E_PREDECESSORMODE = E_PREDECESSORMODE.PM_CREATEWARNINGFLAG
    Private mp_sControlTag As String = ""
    Private mp_oGraphics As DrawingContext
    Private mp_oCulture As System.Globalization.CultureInfo
    Private mp_sStyleIndex As String
    Private mp_oStyle As clsStyle
    Private mp_oImage As Image
    Private mp_sImageTag As String
    Friend mp_oTextBox As clsTextBox
    Public ToolTipEventArgs As ToolTipEventArgs = New ToolTipEventArgs()
    Public ObjectAddedEventArgs As ObjectAddedEventArgs = New ObjectAddedEventArgs()
    Public CustomTierDrawEventArgs As CustomTierDrawEventArgs = New CustomTierDrawEventArgs()
    Public MouseEventArgs As MouseEventArgs = New MouseEventArgs()
    Public MouseHoverEventArgs As MouseEventArgs = New MouseEventArgs()
    Public KeyEventArgs As KeyEventArgs = New KeyEventArgs()
    Public ScrollEventArgs As ScrollEventArgs = New ScrollEventArgs()
    Public DrawEventArgs As DrawEventArgs = New DrawEventArgs()
    Public PredecessorDrawEventArgs As PredecessorDrawEventArgs = New PredecessorDrawEventArgs()
    Public ObjectSelectedEventArgs As ObjectSelectedEventArgs = New ObjectSelectedEventArgs()
    Public ObjectStateChangedEventArgs As ObjectStateChangedEventArgs = New ObjectStateChangedEventArgs()
    Public ErrorEventArgs As ErrorEventArgs = New ErrorEventArgs()
    Public NodeEventArgs As NodeEventArgs = New NodeEventArgs()
    Public MouseWheelEventArgs As MouseWheelEventArgs = New MouseWheelEventArgs()
    Public TextEditEventArgs As TextEditEventArgs = New TextEditEventArgs()
    Public PredecessorExceptionEventArgs As PredecessorExceptionEventArgs = New PredecessorExceptionEventArgs()

    Private Delegate Sub mp_RedrawHandler()

    '//Mouse Events Custom Class
    Public Event ControlClick(ByVal sender As Object, ByVal e As MouseEventArgs)
    Public Event ControlDblClick(ByVal sender As Object, ByVal e As MouseEventArgs)
    Public Event ControlMouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
    Public Event ControlMouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
    Public Event ControlMouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
    Public Event ControlMouseHover(ByVal sender As Object, ByVal e As MouseEventArgs)
    Public Event ControlMouseWheel(ByVal sender As Object, ByVal e As MouseWheelEventArgs)

    '//Keyboard Events
    Public Event ControlKeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
    Public Event ControlKeyPress(ByVal sender As Object, ByVal e As KeyEventArgs)
    Public Event ControlKeyUp(ByVal sender As Object, ByVal e As KeyEventArgs)


    '// Scrolling
    Public Event ControlScroll(ByVal sender As Object, ByVal e As ScrollEventArgs)

    '//Draw
    Public Event Draw(ByVal sender As Object, ByVal e As DrawEventArgs)
    Public Event PredecessorDraw(ByVal sender As Object, ByVal e As PredecessorDrawEventArgs)
    Public Event CustomTierDraw(ByVal sender As Object, ByVal e As CustomTierDrawEventArgs)
    Public Event TierTextDraw(ByVal sender As Object, ByVal e As CustomTierDrawEventArgs)

    '//Added/Selected
    Public Event ObjectAdded(ByVal sender As Object, ByVal e As ObjectAddedEventArgs)
    Public Event ObjectSelected(ByVal sender As Object, ByVal e As ObjectSelectedEventArgs)

    '//Moving
    Public Event BeginObjectMove(ByVal sender As Object, ByVal e As ObjectStateChangedEventArgs)
    Public Event ObjectMove(ByVal sender As Object, ByVal e As ObjectStateChangedEventArgs)
    Public Event EndObjectMove(ByVal sender As Object, ByVal e As ObjectStateChangedEventArgs)
    Public Event CompleteObjectMove(ByVal sender As Object, ByVal e As ObjectStateChangedEventArgs)

    '//Sizing
    Public Event BeginObjectSize(ByVal sender As Object, ByVal e As ObjectStateChangedEventArgs)
    Public Event ObjectSize(ByVal sender As Object, ByVal e As ObjectStateChangedEventArgs)
    Public Event EndObjectSize(ByVal sender As Object, ByVal e As ObjectStateChangedEventArgs)
    Public Event CompleteObjectSize(ByVal sender As Object, ByVal e As ObjectStateChangedEventArgs)

    '//Errors
    Public Event ActiveGanttError(ByVal sender As Object, ByVal e As ErrorEventArgs)
    Public Event PredecessorException(ByVal sender As Object, ByVal e As PredecessorExceptionEventArgs)

    '//Treeview
    Public Event NodeChecked(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeExpanded(ByVal sender As Object, ByVal e As NodeEventArgs)

    '//Text Edit
    Public Event BeginTextEdit(ByVal sender As Object, ByVal e As TextEditEventArgs)
    Public Event EndTextEdit(ByVal sender As Object, ByVal e As TextEditEventArgs)

    Public Event TimeLineChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event ControlRedrawn(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event ControlDraw(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event ToolTipOnMouseHover(ByVal sender As Object, ByVal e As ToolTipEventArgs)
    Public Event ToolTipOnMouseMove(ByVal sender As Object, ByVal e As ToolTipEventArgs)
    Public Event OnMouseHoverToolTipDraw(ByVal sender As Object, ByVal e As ToolTipEventArgs)
    Public Event OnMouseMoveToolTipDraw(ByVal sender As Object, ByVal e As ToolTipEventArgs)

    Friend Sub FirePredecessorException()
        RaiseEvent PredecessorException(Me, PredecessorExceptionEventArgs)
    End Sub

    Friend Sub FireBeginTextEdit()
        RaiseEvent BeginTextEdit(Me, TextEditEventArgs)
    End Sub

    Friend Sub FireEndTextEdit()
        RaiseEvent EndTextEdit(Me, TextEditEventArgs)
    End Sub

    Friend Sub FireControlClick()
        RaiseEvent ControlClick(Me, MouseEventArgs)
    End Sub

    Friend Sub FireControlDblClick()
        RaiseEvent ControlDblClick(Me, MouseEventArgs)
    End Sub

    Friend Sub FireControlMouseDown()
        RaiseEvent ControlMouseDown(Me, MouseEventArgs)
    End Sub

    Friend Sub FireControlMouseMove()
        RaiseEvent ControlMouseMove(Me, MouseEventArgs)
    End Sub

    Friend Sub FireControlMouseUp()
        RaiseEvent ControlMouseUp(Me, MouseEventArgs)
    End Sub

    Friend Sub FireControlMouseHover()
        RaiseEvent ControlMouseHover(Me, MouseHoverEventArgs)
    End Sub

    Friend Sub FireControlMouseWheel()
        RaiseEvent ControlMouseWheel(Me, MouseWheelEventArgs)
    End Sub

    Friend Sub FireControlKeyDown()
        RaiseEvent ControlKeyDown(Me, KeyEventArgs)
    End Sub

    Friend Sub FireControlKeyUp()
        RaiseEvent ControlKeyUp(Me, KeyEventArgs)
    End Sub

    Friend Sub FireControlKeyPress()
        RaiseEvent ControlKeyPress(Me, KeyEventArgs)
    End Sub

    Friend Sub FireDraw()
        RaiseEvent Draw(Me, DrawEventArgs)
    End Sub

    Friend Sub FirePredecessorDraw()
        RaiseEvent PredecessorDraw(Me, PredecessorDrawEventArgs)
    End Sub

    Friend Sub FireCustomTierDraw()
        RaiseEvent CustomTierDraw(Me, CustomTierDrawEventArgs)
    End Sub

    Friend Sub FireTierTextDraw()
        RaiseEvent TierTextDraw(Me, CustomTierDrawEventArgs)
    End Sub

    Friend Sub FireObjectAdded()
        RaiseEvent ObjectAdded(Me, ObjectAddedEventArgs)
    End Sub

    Friend Sub FireObjectSelected()
        RaiseEvent ObjectSelected(Me, ObjectSelectedEventArgs)
    End Sub

    Friend Sub FireBeginObjectMove()
        RaiseEvent BeginObjectMove(Me, ObjectStateChangedEventArgs)
    End Sub

    Friend Sub FireObjectMove()
        RaiseEvent ObjectMove(Me, ObjectStateChangedEventArgs)
    End Sub

    Friend Sub FireEndObjectMove()
        RaiseEvent EndObjectMove(Me, ObjectStateChangedEventArgs)
    End Sub

    Friend Sub FireCompleteObjectMove()
        RaiseEvent CompleteObjectMove(Me, ObjectStateChangedEventArgs)
    End Sub

    Friend Sub FireBeginObjectSize()
        RaiseEvent BeginObjectSize(Me, ObjectStateChangedEventArgs)
    End Sub

    Friend Sub FireObjectSize()
        RaiseEvent ObjectSize(Me, ObjectStateChangedEventArgs)
    End Sub

    Friend Sub FireEndObjectSize()
        RaiseEvent EndObjectSize(Me, ObjectStateChangedEventArgs)
    End Sub

    Friend Sub FireCompleteObjectSize()
        RaiseEvent CompleteObjectSize(Me, ObjectStateChangedEventArgs)
    End Sub

    Friend Sub FireActiveGanttError()
        RaiseEvent ActiveGanttError(Me, ErrorEventArgs)
    End Sub

    Friend Sub FireControlScroll()
        RaiseEvent ControlScroll(Me, ScrollEventArgs)
    End Sub

    Friend Sub FireNodeChecked()
        RaiseEvent NodeChecked(Me, NodeEventArgs)
    End Sub

    Friend Sub FireNodeExpanded()
        RaiseEvent NodeExpanded(Me, NodeEventArgs)
    End Sub

    Friend Sub FireControlDraw()
        RaiseEvent ControlDraw(Me, New System.EventArgs())
    End Sub

    Friend Sub FireControlRedrawn()
        RaiseEvent ControlRedrawn(Me, New System.EventArgs())
    End Sub

    Friend Sub FireTimeLineChanged()
        RaiseEvent TimeLineChanged(Me, New System.EventArgs())
    End Sub

    Friend Sub FireToolTipOnMouseHover(ByVal EventTarget As E_EVENTTARGET)
        If mp_oCurrentView.ClientArea.ToolTipsVisible = False Then
            Return
        End If
        ToolTipEventArgs.ToolTipType = E_TOOLTIPTYPE.TPT_HOVER
        ToolTipEventArgs.EventTarget = EventTarget
        RaiseEvent ToolTipOnMouseHover(Me, ToolTipEventArgs)
    End Sub

    Friend Sub FireToolTipOnMouseMove(ByVal Operation As E_OPERATION)
        If mp_oCurrentView.ClientArea.ToolTipsVisible = False Then
            Return
        End If
        ToolTipEventArgs.ToolTipType = E_TOOLTIPTYPE.TPT_MOVEMENT
        ToolTipEventArgs.Operation = Operation
        RaiseEvent ToolTipOnMouseMove(Me, ToolTipEventArgs)
    End Sub

    Friend Sub FireOnMouseHoverToolTipDraw(ByVal EventTarget As E_EVENTTARGET)
        If mp_oCurrentView.ClientArea.ToolTipsVisible = False Then
            Return
        End If
        ToolTipEventArgs.X1 = ControlToolTip.Left
        ToolTipEventArgs.Y1 = ControlToolTip.Top
        ToolTipEventArgs.X2 = ControlToolTip.Right - 1
        ToolTipEventArgs.Y2 = ControlToolTip.Bottom - 1
        RaiseEvent OnMouseHoverToolTipDraw(Me, ToolTipEventArgs)
    End Sub

    Friend Sub FireOnMouseMoveToolTipDraw(ByVal Operation As E_OPERATION)
        If mp_oCurrentView.ClientArea.ToolTipsVisible = False Then
            Return
        End If
        ToolTipEventArgs.X1 = ControlToolTip.Left
        ToolTipEventArgs.Y1 = ControlToolTip.Top
        ToolTipEventArgs.X2 = ControlToolTip.Right - 1
        ToolTipEventArgs.Y2 = ControlToolTip.Bottom - 1
        RaiseEvent OnMouseMoveToolTipDraw(Me, ToolTipEventArgs)
    End Sub

    Friend Function TempTimeBlocks() As clsTimeBlocks
        Return tmpTimeBlocks
    End Function

    Public Sub New()
        InitializeComponent()
        RenderOptions.SetEdgeMode(Me, EdgeMode.Aliased)
        mp_oLicense = LicenseManager.Validate(GetType(ActiveGanttVBWCtl), Me)

#If DemoVersion Then
        If ShowAbout() = True Then
            AboutBox()
        End If
#End If

        mp_bPropertiesRead = False
        If (mp_bPropertiesRead = False) Then
            clsG = New clsGraphics(Me)
            MathLib = New clsMath(Me)
            StrLib = New clsString(Me)
            Styles = New clsStyles(Me)
            mp_sStyleIndex = "DS_CONTROL"
            mp_oStyle = Styles.FItem("DS_CONTROL")
            VerticalScrollBar = New clsVerticalScrollBar(Me)
            HorizontalScrollBar = New clsHorizontalScrollBar(Me)
            Rows = New clsRows(Me)
            Tasks = New clsTasks(Me)
            Columns = New clsColumns(Me)
            Layers = New clsLayers(Me)
            Percentages = New clsPercentages(Me)
            TimeBlocks = New clsTimeBlocks(Me)
            Predecessors = New clsPredecessors(Me)
            tmpTimeBlocks = New clsTimeBlocks(Me)
            Printer = New clsPrinter(Me)
            Splitter = New clsSplitter(Me)
            Views = New clsViews(Me)
            Treeview = New clsTreeview(Me)
            mp_oCurrentView = Views.FItem("0")
            MouseKeyboardEvents = New clsMouseKeyboardEvents(Me)
            Drawing = New clsDrawing(Me)
            mp_oCulture = System.Globalization.CultureInfo.CurrentCulture.Clone()
            TierAppearance = New clsTierAppearance(Me)
            TierFormat = New clsTierFormat(Me)
            ScrollBarSeparator = New clsScrollBarSeparator(Me)
            mp_oTextBox = New clsTextBox(Me)


            mp_oImage = Nothing
            mp_sImageTag = ""
        End If
        mp_bPropertiesRead = True
        Me.Focusable = True
    End Sub

    Private Sub ActiveGanttVBWCtl_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        mp_oLicense = LicenseManager.Validate(GetType(ActiveGanttVBWCtl), Me)
        Me.Focus()
    End Sub

    Private Sub ActiveGanttVBWCtl_Unloaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Unloaded
        If Not (mp_oLicense Is Nothing) Then
            Me.mp_oLicense.Dispose()
            Me.mp_oLicense = Nothing
        End If
    End Sub

    Friend Sub ForceRender()
        Dim oDrawingVisual As New DrawingVisual()
        Dim oDC As DrawingContext = oDrawingVisual.RenderOpen()
        OnRender(oDC)
        oDC.Close()
    End Sub

    Protected Overrides Sub OnRender(ByVal oDC As DrawingContext)
        If mp_bPropertiesRead = False Then
            Return
        End If
        mp_oGraphics = oDC
        If System.ComponentModel.DesignerProperties.GetIsInDesignMode(Me) = True Then
            mp_DrawDesignMode()
        Else
            'mp_CHKPXPScrollButtons()
            'mp_CHKPXPLines()
            'mp_CHKPXPButtons()
            'mp_CHKPXPArrows()
            'mp_CHKPXPFigures()
            'mp_CHKPXPGradients()
            'mp_CHKPXPText()
            'mp_CHKPXPHatch()
            mp_Draw()
        End If
        If mp_oCurrentView.TimeLine.StartDate <> mp_dtTimeLineStartBuffer Or mp_oCurrentView.TimeLine.EndDate <> mp_dtTimeLineEndBuffer Then
            mp_dtTimeLineStartBuffer = mp_oCurrentView.TimeLine.StartDate
            mp_dtTimeLineEndBuffer = mp_oCurrentView.TimeLine.EndDate
            FireTimeLineChanged()
        End If
        MyBase.OnRender(oDC)
    End Sub

    Private Sub ActiveGanttVBWCtl_SizeChanged(ByVal sender As Object, ByVal e As System.Windows.SizeChangedEventArgs) Handles Me.SizeChanged
        If mp_bPropertiesRead = False Then
            Return
        End If
        Redraw()
    End Sub

    Private Sub mp_Draw()
        clsG.RequiresPop = False
        FireControlDraw()
        clsG.ResetFocusRectangle()
        clsG.ClipRegion(0, 0, clsG.Width, clsG.Height, False)
        clsG.mp_DrawItem(0, 0, clsG.Width - 1, clsG.Height - 1, "", "", False, Me.Image, 0, 0, Me.Style)
        mp_oCurrentView.TimeLine.Calculate()
        mp_PositionScrollBars()
        Columns.Position()
        Rows.InitializePosition()
        Rows.PositionRows()
        mp_oCurrentView.TimeLine.Draw()
        mp_oCurrentView.TimeLine.ProgressLine.Draw()
            Columns.Draw()
            Rows.Draw()
            Treeview.Draw()
        TimeBlocks.CreateTemporaryTimeBlocks()
        TimeBlocks.Draw()
        mp_oCurrentView.ClientArea.Grid.Draw()
        mp_oCurrentView.ClientArea.Draw()
        Predecessors.Draw()
        Tasks.Draw()
        Percentages.Draw()
        mp_oCurrentView.TimeLine.ProgressLine.Draw()
        Splitter.Draw()
        clsG.ClipRegion(0, 0, clsG.Width, clsG.Height, False)
        If VerticalScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            clsG.mp_DrawItem(VerticalScrollBar.Left, VerticalScrollBar.Top + VerticalScrollBar.Height, VerticalScrollBar.Left + 16, VerticalScrollBar.Top + VerticalScrollBar.Height + 16, "", "", False, Nothing, 0, 0, ScrollBarSeparator.Style)
            clsG.ClipRegion(0, 0, clsG.Width, clsG.Height, False)
        ElseIf mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            clsG.mp_DrawItem(mp_oCurrentView.TimeLine.TimeLineScrollBar.Left + mp_oCurrentView.TimeLine.TimeLineScrollBar.Width, mp_oCurrentView.TimeLine.TimeLineScrollBar.Top, mp_oCurrentView.TimeLine.TimeLineScrollBar.Left + mp_oCurrentView.TimeLine.TimeLineScrollBar.Width + 16, mp_oCurrentView.TimeLine.TimeLineScrollBar.Top + 16, "", "", False, Nothing, 0, 0, ScrollBarSeparator.Style)
            clsG.ClipRegion(0, 0, clsG.Width, clsG.Height, False)
        End If
        mp_DrawDebugMetrics()
        If VerticalScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            VerticalScrollBar.ScrollBar.Draw()

        End If
        If HorizontalScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            HorizontalScrollBar.ScrollBar.Draw()
        End If
        If mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            mp_oCurrentView.TimeLine.TimeLineScrollBar.ScrollBar.Draw()
        End If
#If DemoVersion Then
        Dim oFont As New Font("Arial", 12, System.Windows.FontWeights.Bold)
        Dim rnd As System.Random
        rnd = New System.Random()
        Dim oColor As System.Windows.Media.Color = New System.Windows.Media.Color()
        oColor = Color.FromRgb(rnd.Next(0, 255), rnd.Next(0, 255), rnd.Next(0, 255))
        clsG.DrawAlignedText(20, 20, clsG.Width() - 30, clsG.Height() - 20, "ActiveGanttVBW Scheduler Component" & vbCrLf & "Trial Version: " & Version & vbCrLf & "For evaluation purposes only" & vbCrLf & "Purchase the full version through: " & vbCrLf & "http://www.sourcecodestore.com", GRE_HORIZONTALALIGNMENT.HAL_RIGHT, GRE_VERTICALALIGNMENT.VAL_BOTTOM, oColor, oFont, True)
#End If
        FireControlRedrawn()

    End Sub

    Private Sub mp_DrawDesignMode()
        Dim lLeftBox As Integer
        Dim lTop As Integer
        Dim lRightBox As Integer
        Dim lBottom As Integer
        Dim lLeftCA As Integer
        Dim lRightCA As Integer
        clsG.RequiresPop = False
        Dim oFont As Font = New Font("Arial", 8)
        clsG.DrawLine(0, 0, clsG.Width, clsG.Height, GRE_LINETYPE.LT_FILLED, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
        mp_oCurrentView.TimeLine.Calculate()
        mp_oCurrentView.TimeLine.Draw()
        clsG.ClearClipRegion()
        lLeftBox = mt_LeftMargin
        lTop = mt_TopMargin
        lRightBox = Splitter.Left
        lBottom = mp_oCurrentView.TimeLine.Bottom
        clsG.DrawEdge(lLeftBox, lTop, lRightBox, lBottom, System.Windows.Media.Color.FromRgb(192, 192, 192), GRE_BUTTONSTYLE.BT_NORMALWINDOWS, GRE_EDGETYPE.ET_RAISED, True, Nothing)
        clsG.DrawAlignedText(lLeftBox, lTop, lRightBox, lBottom, "Column", GRE_HORIZONTALALIGNMENT.HAL_CENTER, GRE_VERTICALALIGNMENT.VAL_CENTER, System.Windows.Media.Colors.Black, oFont)
        lLeftBox = mt_LeftMargin
        lTop = mp_oCurrentView.ClientArea.Top
        lRightBox = Splitter.Left
        lBottom = mp_oCurrentView.ClientArea.Top + 40
        lLeftCA = Splitter.Right
        lRightCA = mt_RightMargin
        clsG.DrawEdge(lLeftBox, lTop, lRightBox, lBottom, System.Windows.Media.Color.FromRgb(192, 192, 192), GRE_BUTTONSTYLE.BT_NORMALWINDOWS, GRE_EDGETYPE.ET_RAISED, True, Nothing)
        clsG.DrawAlignedText(lLeftBox, lTop, lRightBox, lBottom, "Cell", GRE_HORIZONTALALIGNMENT.HAL_CENTER, GRE_VERTICALALIGNMENT.VAL_CENTER, System.Windows.Media.Colors.Black, oFont)
        If mp_oCurrentView.ClientArea.Grid.HorizontalLines = True Then
            clsG.DrawLine(lLeftCA, lBottom, lRightCA, lBottom, GRE_LINETYPE.LT_NORMAL, mp_oCurrentView.ClientArea.Grid.Color, GRE_LINEDRAWSTYLE.LDS_SOLID)
        End If
        Rows.TopOffset = mp_oCurrentView.ClientArea.Top + 40
        mp_oCurrentView.ClientArea.f_LastVisibleRow = 1
        Splitter.Draw()
        clsG.ClipRegion(0, 0, clsG.Width, clsG.Height, False)
        clsG.DrawEdge(0, 0, clsG.Width - 1, clsG.Height - 1, Colors.Black, GRE_BUTTONSTYLE.BT_NORMALWINDOWS, GRE_EDGETYPE.ET_SUNKEN, False, Nothing)
    End Sub

    Private Sub mp_DrawDebugMetrics()

    End Sub

    Friend Function f_HDC() As DrawingContext
        Return mp_oGraphics
    End Function

    Friend Function f_Canvas() As Canvas
        Return oCanvas
    End Function

    Friend Function f_Width() As Integer
        Return Me.Width
    End Function

    Friend Function f_Height() As Integer
        Return Me.Height
    End Function

    Friend Function mp_lStrWidth(ByVal sString As String, ByVal r_oFont As Font) As Integer
        Dim oTypeFace As New Typeface(r_oFont.FamilyName)
        Dim oFormattedText As New FormattedText(sString, Culture, FlowDirection.LeftToRight, oTypeFace, r_oFont.WPFFontSize, New SolidColorBrush(Colors.Black))
        Return System.Convert.ToInt32(oFormattedText.Width())
    End Function

    Friend Function mp_lStrHeight(ByVal sString As String, ByVal r_oFont As Font) As Integer
        Dim oTypeFace As New Typeface(r_oFont.FamilyName)
        Dim oFormattedText As New FormattedText(sString, Culture, FlowDirection.LeftToRight, oTypeFace, r_oFont.WPFFontSize, New SolidColorBrush(Colors.Black))
        Return System.Convert.ToInt32(oFormattedText.Height())
    End Function

    Friend Function GetGraphicsObject() As Canvas
        Return oCanvas
    End Function

    Friend Sub f_Draw()
        mp_Draw()
    End Sub

    Friend ReadOnly Property f_UserMode() As Boolean
        Get
            Return Not System.ComponentModel.DesignerProperties.GetIsInDesignMode(Me)
        End Get
    End Property

    Friend ReadOnly Property mt_BorderThickness() As Integer
        Get
            Select Case mp_oStyle.Appearance
                Case E_STYLEAPPEARANCE.SA_RAISED
                    Return 2
                Case E_STYLEAPPEARANCE.SA_SUNKEN
                    Return 2
                Case E_STYLEAPPEARANCE.SA_FLAT
                    If mp_oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_NONE Then
                        Return 0
                    Else
                        Return mp_oStyle.BorderWidth
                    End If
                Case E_STYLEAPPEARANCE.SA_CELL
                    If mp_oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_NONE Then
                        Return 0
                    Else
                        Return mp_oStyle.BorderWidth
                    End If
                Case E_STYLEAPPEARANCE.SA_GRAPHICAL
                    Return 0
            End Select
            Return 0
        End Get
    End Property

    Friend ReadOnly Property mt_TableBottom() As Integer
        Get
            If HorizontalScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
                Return clsG.Height - mt_BorderThickness - 1 - HorizontalScrollBar.Height
            Else
                Return clsG.Height - mt_BorderThickness - 1
            End If
        End Get
    End Property

    Friend ReadOnly Property mt_TopMargin() As Integer
        Get
            Return mt_BorderThickness
        End Get
    End Property

    Friend ReadOnly Property mt_LeftMargin() As Integer
        Get
            Return mt_BorderThickness
        End Get
    End Property

    Friend ReadOnly Property mt_RightMargin() As Integer
        Get
            If VerticalScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
                Return clsG.Width - mt_BorderThickness - 1 - VerticalScrollBar.Width
            Else
                Return clsG.Width - mt_BorderThickness - 1
            End If
        End Get
    End Property

    Friend ReadOnly Property mt_BottomMargin() As Integer
        Get
            Return clsG.Height - mt_BorderThickness - 1
        End Get
    End Property

    Private Sub ActiveGanttVBWCtl_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles Me.MouseLeave
        If System.ComponentModel.DesignerProperties.GetIsInDesignMode(Me) = True Then Return
        MouseKeyboardEvents.OnMouseLeave()
    End Sub

    Private Sub ActiveGanttVBWCtl_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Me.MouseDown
        If System.ComponentModel.DesignerProperties.GetIsInDesignMode(Me) = True Then Return
        MouseKeyboardEvents.OnMouseDown(e.GetPosition(oCanvas).X, e.GetPosition(oCanvas).Y, e.ChangedButton())
    End Sub

    Private Sub ActiveGanttVBWCtl_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles Me.MouseMove
        If System.ComponentModel.DesignerProperties.GetIsInDesignMode(Me) = True Then Return
        MouseKeyboardEvents.OnMouseMoveGeneral(e.GetPosition(oCanvas).X, e.GetPosition(oCanvas).Y)
    End Sub

    Private Sub ActiveGanttVBWCtl_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Me.MouseUp
        If System.ComponentModel.DesignerProperties.GetIsInDesignMode(Me) = True Then Return
        MouseKeyboardEvents.OnMouseUp(e.GetPosition(oCanvas).X, e.GetPosition(oCanvas).Y)
        Redraw()
    End Sub

    Private Sub ActiveGanttVBWCtl_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles Me.KeyDown
        If System.ComponentModel.DesignerProperties.GetIsInDesignMode(Me) = True Then Return
        MouseKeyboardEvents.KeyDown(e.Key)
    End Sub

    Private Sub ActiveGanttVBWCtl_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles Me.KeyUp
        If System.ComponentModel.DesignerProperties.GetIsInDesignMode(Me) = True Then Return
        MouseKeyboardEvents.KeyUp(e.Key)
    End Sub

    Private Sub ActiveGanttVBWCtl_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Input.MouseWheelEventArgs) Handles Me.MouseWheel
        If System.ComponentModel.DesignerProperties.GetIsInDesignMode(Me) = True Then Return
        MouseWheelEventArgs.Clear()
        MouseWheelEventArgs.Delta = e.Delta
        MouseWheelEventArgs.X = System.Convert.ToInt32(e.GetPosition(Me).X)
        MouseWheelEventArgs.Y = System.Convert.ToInt32(e.GetPosition(Me).Y)
        If e.LeftButton = MouseButtonState.Pressed Then
            MouseWheelEventArgs.Button = E_MOUSEBUTTONS.BTN_LEFT
        ElseIf e.RightButton = MouseButtonState.Pressed Then
            MouseWheelEventArgs.Button = E_MOUSEBUTTONS.BTN_RIGHT
        ElseIf e.MiddleButton = MouseButtonState.Pressed Then
            MouseWheelEventArgs.Button = E_MOUSEBUTTONS.BTN_MIDDLE
        ElseIf e.XButton1 = MouseButtonState.Pressed Then
            MouseWheelEventArgs.Button = E_MOUSEBUTTONS.BTN_XBUTTON1
        ElseIf e.XButton2 = MouseButtonState.Pressed Then
            MouseWheelEventArgs.Button = E_MOUSEBUTTONS.BTN_XBUTTON2
        Else
            MouseWheelEventArgs.Button = E_MOUSEBUTTONS.BTN_NONE
        End If
        FireControlMouseWheel()

    End Sub

    Friend Sub VerticalScrollBar_ValueChanged(ByVal Offset As Integer)
        Redraw()
        ScrollEventArgs.Clear()
        ScrollEventArgs.ScrollBarType = E_SCROLLBAR.SCR_VERTICAL
        ScrollEventArgs.Offset = Offset
        FireControlScroll()
    End Sub

    Friend Sub HorizontalScrollBar_ValueChanged(ByVal Offset As Integer)
        Redraw()
        ScrollEventArgs.Clear()
        ScrollEventArgs.ScrollBarType = E_SCROLLBAR.SCR_HORIZONTAL1
        ScrollEventArgs.Offset = Offset
        FireControlScroll()
    End Sub

    Friend Sub TimeLineScrollBar_ValueChanged(ByVal Offset As Integer)
        Redraw()
        ScrollEventArgs.Clear()
        ScrollEventArgs.ScrollBarType = E_SCROLLBAR.SCR_HORIZONTAL2
        ScrollEventArgs.Offset = Offset
        FireControlScroll()
    End Sub

    Friend Sub mp_PositionScrollBars()
        If System.ComponentModel.DesignerProperties.GetIsInDesignMode(Me) = True Or clsG.CustomPrinting = True Then
            VerticalScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
            HorizontalScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
            mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
            Return
        End If

        If clsG.Height <= mp_oCurrentView.ClientArea.Top Then
            VerticalScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
            HorizontalScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
            mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
            Return
        End If

        '// Determine need for HorizontalScrollBar
        Dim lWidth As Integer
        lWidth = Columns.Width
        If lWidth > Splitter.Right Then
            If HorizontalScrollBar.mf_Visible = True Then
                HorizontalScrollBar.State = E_SCROLLSTATE.SS_NEEDED
            Else
                HorizontalScrollBar.State = E_SCROLLSTATE.SS_NOTNEEDED
            End If
        Else
            HorizontalScrollBar.State = E_SCROLLSTATE.SS_NOTNEEDED
        End If
        If Splitter.Right < 5 Then
            HorizontalScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
        End If

        '// Determine need for mp_oCurrentView.TimeLine.TimeLineScrollBar
        If Splitter.Right < clsG.Width - (18 + mt_BorderThickness) Then
            If mp_oCurrentView.TimeLine.TimeLineScrollBar.Enabled = True Then
                If mp_oCurrentView.TimeLine.TimeLineScrollBar.mf_Visible = True Then
                    mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_NEEDED
                Else
                    mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_NOTNEEDED
                End If
            Else
                mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_NOTNEEDED
            End If
        Else
            mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
        End If

        '// Determine need for VerticalScrollBar
        If ((Rows.Height() + mp_oCurrentView.ClientArea.Top + HorizontalScrollBar.Height + mt_BorderThickness) > clsG.Height) Or (Rows.RealFirstVisibleRow > 1) Then
            If mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY Then
                VerticalScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
            Else
                VerticalScrollBar.State = E_SCROLLSTATE.SS_NEEDED
            End If
        Else
            VerticalScrollBar.State = E_SCROLLSTATE.SS_NOTNEEDED
        End If

        If VerticalScrollBar.mf_Visible = False Then
            VerticalScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
        End If
        If HorizontalScrollBar.mf_Visible = False Then
            HorizontalScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
        End If
        If mp_oCurrentView.TimeLine.TimeLineScrollBar.mf_Visible = False Then
            mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
        End If

        If VerticalScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            VerticalScrollBar.Position()
        End If
        If HorizontalScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            HorizontalScrollBar.Position()
        End If
        If mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            mp_oCurrentView.TimeLine.TimeLineScrollBar.Position()
        End If
    End Sub

    Public Sub WriteXML(ByVal url As String)
        Dim oXML As New clsXML(Me, "ActiveGanttCtl")
        mp_WriteXML(oXML)
        oXML.WriteXML(url)
    End Sub

    Public Sub ReadXML(ByVal url As String)
        Dim oXML As New clsXML(Me, "ActiveGanttCtl")
        oXML.ReadXML(url)
        mp_ReadXML(oXML)
    End Sub

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(Me, "ActiveGanttCtl")
        oXML.SetXML(sXML)
        mp_ReadXML(oXML)
    End Sub

    Public Function GetXML() As String
        Dim oXML As New clsXML(Me, "ActiveGanttCtl")
        mp_WriteXML(oXML)
        Return oXML.GetXML
    End Function

    Private Sub mp_WriteXML(ByRef oXML As clsXML)
        oXML.InitializeWriter()
        oXML.WriteProperty("Version", "AGVBW")
        oXML.WriteProperty("ControlTag", mp_sControlTag)
        oXML.WriteProperty("AddMode", mp_yAddMode)
        oXML.WriteProperty("AddDurationInterval", mp_yAddDurationInterval)
        oXML.WriteProperty("AllowAdd", mp_bAllowAdd)
        oXML.WriteProperty("AllowColumnMove", mp_bAllowColumnMove)
        oXML.WriteProperty("AllowColumnSize", mp_bAllowColumnSize)
        oXML.WriteProperty("AllowEdit", mp_bAllowEdit)
        oXML.WriteProperty("AllowPredecessorAdd", mp_bAllowPredecessorAdd)
        oXML.WriteProperty("AllowRowMove", mp_bAllowRowMove)
        oXML.WriteProperty("AllowRowSize", mp_bAllowRowSize)
        oXML.WriteProperty("AllowSplitterMove", mp_bAllowSplitterMove)
        oXML.WriteProperty("AllowTimeLineScroll", mp_bAllowTimeLineScroll)
        oXML.WriteProperty("EnforcePredecessors", mp_bEnforcePredecessors)
        oXML.WriteProperty("CurrentLayer", mp_sCurrentLayer)
        oXML.WriteProperty("CurrentView", mp_sCurrentView)
        oXML.WriteProperty("DoubleBuffering", mp_bDoubleBuffering)
        oXML.WriteProperty("ErrorReports", mp_yErrorReports)
        oXML.WriteProperty("LayerEnableObjects", mp_yLayerEnableObjects)
        oXML.WriteProperty("MinColumnWidth", mp_lMinColumnWidth)
        oXML.WriteProperty("MinRowHeight", mp_lMinRowHeight)
        oXML.WriteProperty("ScrollBarBehaviour", mp_yScrollBarBehaviour)
        oXML.WriteProperty("SelectedCellIndex", mp_lSelectedCellIndex)
        oXML.WriteProperty("SelectedColumnIndex", mp_lSelectedColumnIndex)
        oXML.WriteProperty("SelectedPercentageIndex", mp_lSelectedPercentageIndex)
        oXML.WriteProperty("SelectedPredecessorIndex", mp_lSelectedPredecessorIndex)
        oXML.WriteProperty("SelectedRowIndex", mp_lSelectedRowIndex)
        oXML.WriteProperty("SelectedTaskIndex", mp_lSelectedTaskIndex)
        oXML.WriteProperty("TreeviewColumnIndex", mp_lTreeviewColumnIndex)
        oXML.WriteProperty("TimeBlockBehaviour", mp_yTimeBlockBehaviour)
        oXML.WriteProperty("TierAppearanceScope", mp_yTierAppearanceScope)
        oXML.WriteProperty("TierFormatScope", mp_yTierFormatScope)
        oXML.WriteProperty("PredecessorMode", mp_yPredecessorMode)
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        oXML.WriteProperty("Image", mp_oImage)
        oXML.WriteProperty("ImageTag", mp_sImageTag)
        oXML.WriteObject(Styles.GetXML())
        oXML.WriteObject(Rows.GetXML())
        oXML.WriteObject(Columns.GetXML())
        oXML.WriteObject(Layers.GetXML())
        oXML.WriteObject(Tasks.GetXML())
        oXML.WriteObject(Predecessors.GetXML())
        oXML.WriteObject(Views.GetXML())
        oXML.WriteObject(TimeBlocks.GetXML())
        oXML.WriteObject(TimeBlocks.CP_GetXML())
        oXML.WriteObject(Percentages.GetXML())
        oXML.WriteObject(Splitter.GetXML())
        oXML.WriteObject(Treeview.GetXML())
        oXML.WriteObject(TierAppearance.GetXML())
        oXML.WriteObject(TierFormat.GetXML())
        oXML.WriteObject(ScrollBarSeparator.GetXML())
        oXML.WriteObject(VerticalScrollBar.GetXML())
        oXML.WriteObject(HorizontalScrollBar.GetXML())
    End Sub

    Private Sub mp_ReadXML(ByRef oXML As clsXML)
        Dim sVersion As String = ""
        Dim sCurrentView As String = ""
        Clear()
        oXML.InitializeReader()
        oXML.ReadProperty("Version", sVersion)
        oXML.ReadProperty("ControlTag", mp_sControlTag)
        oXML.ReadProperty("AddMode", mp_yAddMode)
        oXML.ReadProperty("AddDurationInterval", mp_yAddDurationInterval)
        oXML.ReadProperty("AllowAdd", mp_bAllowAdd)
        oXML.ReadProperty("AllowColumnMove", mp_bAllowColumnMove)
        oXML.ReadProperty("AllowColumnSize", mp_bAllowColumnSize)
        oXML.ReadProperty("AllowEdit", mp_bAllowEdit)
        oXML.ReadProperty("AllowPredecessorAdd", mp_bAllowPredecessorAdd)
        oXML.ReadProperty("AllowRowMove", mp_bAllowRowMove)
        oXML.ReadProperty("AllowRowSize", mp_bAllowRowSize)
        oXML.ReadProperty("AllowSplitterMove", mp_bAllowSplitterMove)
        oXML.ReadProperty("AllowTimeLineScroll", mp_bAllowTimeLineScroll)
        oXML.ReadProperty("EnforcePredecessors", mp_bEnforcePredecessors)
        oXML.ReadProperty("CurrentLayer", mp_sCurrentLayer)
        oXML.ReadProperty("CurrentView", mp_sCurrentView)
        oXML.ReadProperty("DoubleBuffering", mp_bDoubleBuffering)
        oXML.ReadProperty("ErrorReports", mp_yErrorReports)
        oXML.ReadProperty("LayerEnableObjects", mp_yLayerEnableObjects)
        oXML.ReadProperty("MinColumnWidth", mp_lMinColumnWidth)
        oXML.ReadProperty("MinRowHeight", mp_lMinRowHeight)
        oXML.ReadProperty("ScrollBarBehaviour", mp_yScrollBarBehaviour)
        oXML.ReadProperty("SelectedCellIndex", mp_lSelectedCellIndex)
        oXML.ReadProperty("SelectedColumnIndex", mp_lSelectedColumnIndex)
        oXML.ReadProperty("SelectedPercentageIndex", mp_lSelectedPercentageIndex)
        oXML.ReadProperty("SelectedPredecessorIndex", mp_lSelectedPredecessorIndex)
        oXML.ReadProperty("SelectedRowIndex", mp_lSelectedRowIndex)
        oXML.ReadProperty("SelectedTaskIndex", mp_lSelectedTaskIndex)
        oXML.ReadProperty("TreeviewColumnIndex", mp_lTreeviewColumnIndex)
        oXML.ReadProperty("TimeBlockBehaviour", mp_yTimeBlockBehaviour)
        oXML.ReadProperty("TierAppearanceScope", mp_yTierAppearanceScope)
        oXML.ReadProperty("TierFormatScope", mp_yTierFormatScope)
        oXML.ReadProperty("PredecessorMode", mp_yPredecessorMode)
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        oXML.ReadProperty("Image", mp_oImage)
        oXML.ReadProperty("ImageTag", mp_sImageTag)
        Styles.SetXML(oXML.ReadObject("Styles"))
        Rows.SetXML(oXML.ReadObject("Rows"))
        Columns.SetXML(oXML.ReadObject("Columns"))
        Layers.SetXML(oXML.ReadObject("Layers"))
        Tasks.SetXML(oXML.ReadObject("Tasks"))
        Predecessors.SetXML(oXML.ReadObject("Predecessors"))
        Views.SetXML(oXML.ReadObject("Views"))
        TimeBlocks.SetXML(oXML.ReadObject("TimeBlocks"))
        TimeBlocks.CP_SetXML(oXML.ReadObject("CP_TimeBlocks"))
        Percentages.SetXML(oXML.ReadObject("Percentages"))
        Splitter.SetXML(oXML.ReadObject("Splitter"))
        Treeview.SetXML(oXML.ReadObject("Treeview"))
        TierAppearance.SetXML(oXML.ReadObject("TierAppearance"))
        TierFormat.SetXML(oXML.ReadObject("TierFormat"))
        ScrollBarSeparator.SetXML(oXML.ReadObject("ScrollBarSeparator"))
        VerticalScrollBar.SetXML(oXML.ReadObject("VerticalScrollBar"))
        HorizontalScrollBar.SetXML(oXML.ReadObject("HorizontalScrollBar"))
        StyleIndex = mp_sStyleIndex
        Rows.UpdateTree()
        CurrentView = mp_sCurrentView
        mp_oCurrentView.TimeLine.Position(mp_oCurrentView.TimeLine.StartDate)
    End Sub

    Friend Sub mp_ErrorReport(ByVal ErrNumber As Integer, ByVal ErrDescription As String, ByVal ErrSource As String)
        If mp_yErrorReports = E_REPORTERRORS.RE_MSGBOX Then
            MessageBox.Show(System.Convert.ToString(ErrNumber) & ": " & ErrDescription & " (" & ErrSource & ")", "AGVBW Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error)
        ElseIf mp_yErrorReports = E_REPORTERRORS.RE_HIDE Then
        ElseIf mp_yErrorReports = E_REPORTERRORS.RE_RAISE Then
            Dim ex As AGError = New AGError(ErrNumber.ToString() & ": " & ErrDescription + " - " & ErrSource)
            ex.ErrNumber = ErrNumber
            ex.ErrDescription = ErrDescription
            ex.ErrSource = ErrSource
            Throw ex
        ElseIf mp_yErrorReports = E_REPORTERRORS.RE_RAISEEVENT Then
            ErrorEventArgs.Clear()
            ErrorEventArgs.Number = ErrNumber
            ErrorEventArgs.Description = ErrDescription
            ErrorEventArgs.Source = ErrSource
            FireActiveGanttError()
        End If
    End Sub

    Public Property ErrorReports() As E_REPORTERRORS
        Get
            Return mp_yErrorReports
        End Get
        Set(ByVal Value As E_REPORTERRORS)
            mp_yErrorReports = Value
        End Set
    End Property

    Public Property CurrentLayer() As String
        Get
            Return mp_sCurrentLayer
        End Get
        Set(ByVal Value As String)
            mp_sCurrentLayer = Value
        End Set
    End Property

    Public Property CurrentView() As String
        Get
            Return mp_sCurrentView
        End Get
        Set(ByVal Value As String)
            If Value = "" Then
                Value = "0"
            End If
            mp_oCurrentView = Views.FItem(Value)
            mp_sCurrentView = Value
        End Set
    End Property

    Public ReadOnly Property CurrentViewObject() As clsView
        Get
            Return mp_oCurrentView
        End Get
    End Property

    Public ReadOnly Property ControlToolTip() As clsToolTip
        Get
            Return MouseKeyboardEvents.mp_oToolTip
        End Get
    End Property

    Public Property ScrollBarBehaviour() As E_SCROLLBEHAVIOUR
        Get
            Return mp_yScrollBarBehaviour
        End Get
        Set(ByVal Value As E_SCROLLBEHAVIOUR)
            mp_yScrollBarBehaviour = Value
        End Set
    End Property

    Public Property TierAppearanceScope() As E_TIERAPPEARANCESCOPE
        Get
            Return mp_yTierAppearanceScope
        End Get
        Set(ByVal Value As E_TIERAPPEARANCESCOPE)
            mp_yTierAppearanceScope = Value
        End Set
    End Property

    Public Property TierFormatScope() As E_TIERFORMATSCOPE
        Get
            Return mp_yTierFormatScope
        End Get
        Set(ByVal Value As E_TIERFORMATSCOPE)
            mp_yTierFormatScope = Value
        End Set
    End Property

    Public Property TimeBlockBehaviour() As E_TIMEBLOCKBEHAVIOUR
        Get
            Return mp_yTimeBlockBehaviour
        End Get
        Set(ByVal Value As E_TIMEBLOCKBEHAVIOUR)
            mp_yTimeBlockBehaviour = Value
        End Set
    End Property

    Public Property SelectedTaskIndex() As Integer
        Get
            Return mp_lSelectedTaskIndex
        End Get
        Set(ByVal Value As Integer)
            If Value <= 0 Then
                Value = 0
            ElseIf Value > Tasks.Count Then
                Value = Tasks.Count
            End If
            mp_lSelectedTaskIndex = Value
        End Set
    End Property

    Public Property SelectedColumnIndex() As Integer
        Get
            Return mp_lSelectedColumnIndex
        End Get
        Set(ByVal Value As Integer)
            If Value <= 0 Then
                Value = 0
            ElseIf Value > Columns.Count Then
                Value = Columns.Count
            End If
            mp_lSelectedColumnIndex = Value
        End Set
    End Property

    Public Property SelectedRowIndex() As Integer
        Get
            Return mp_lSelectedRowIndex
        End Get
        Set(ByVal Value As Integer)
            If Value <= 0 Then
                Value = 0
            ElseIf Value > Rows.Count Then
                Value = Rows.Count
            End If
            mp_lSelectedRowIndex = Value
        End Set
    End Property

    Public Property SelectedCellIndex() As Integer
        Get
            Return mp_lSelectedCellIndex
        End Get
        Set(ByVal Value As Integer)
            If Value <= 0 Then
                Value = 0
            ElseIf Value > Columns.Count Then
                Value = Columns.Count
            End If
            mp_lSelectedCellIndex = Value
        End Set
    End Property

    Public Property SelectedPercentageIndex() As Integer
        Get
            Return mp_lSelectedPercentageIndex
        End Get
        Set(ByVal Value As Integer)
            If Value <= 0 Then
                Value = 0
            ElseIf Value > Percentages.Count Then
                Value = Percentages.Count
            End If
            mp_lSelectedPercentageIndex = Value
        End Set
    End Property

    Public Property SelectedPredecessorIndex() As Integer
        Get
            Return mp_lSelectedPredecessorIndex
        End Get
        Set(ByVal value As Integer)
            If value <= 0 Then
                value = 0
            ElseIf value > Percentages.Count Then
                value = Percentages.Count
            End If
            mp_lSelectedPredecessorIndex = value
        End Set
    End Property

    Public Property TreeviewColumnIndex() As Integer
        Get
            If Columns.Count = 0 Then
                Return 0
            ElseIf mp_lTreeviewColumnIndex > Columns.Count Then
                Return 0
            ElseIf mp_lTreeviewColumnIndex < 0 Then
                Return 0
            Else
                Return mp_lTreeviewColumnIndex
            End If
        End Get
        Set(ByVal value As Integer)
            If value <= 0 Then
                value = 0
            ElseIf value > Columns.Count Then
                value = Columns.Count
            End If
            mp_lTreeviewColumnIndex = value
        End Set
    End Property

    Public Property StyleIndex() As String
        Get
            If mp_sStyleIndex = "DS_CONTROL" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_CONTROL"
            mp_sStyleIndex = Value
            mp_oStyle = Styles.FItem(Value)
        End Set
    End Property

    Public Shadows ReadOnly Property Style() As clsStyle
        Get
            Return mp_oStyle
        End Get
    End Property

    Public Property Image() As Image
        Get
            Return mp_oImage
        End Get
        Set(ByVal Value As Image)
            mp_oImage = Value
        End Set
    End Property

    Public Property ImageTag() As String
        Get
            Return mp_sImageTag
        End Get
        Set(ByVal Value As String)
            mp_sImageTag = Value
        End Set
    End Property

    Public Property Culture() As System.Globalization.CultureInfo
        Get
            Return mp_oCulture
        End Get
        Set(ByVal Value As System.Globalization.CultureInfo)
            mp_oCulture = Value
        End Set
    End Property

    Public Sub ClearSelections()
        mp_oTextBox.Terminate()
        mp_lSelectedTaskIndex = 0
        mp_lSelectedColumnIndex = 0
        mp_lSelectedRowIndex = 0
        mp_lSelectedCellIndex = 0
        mp_lSelectedPercentageIndex = 0
        mp_lSelectedPredecessorIndex = 0
    End Sub

    Public Sub Clear()
        mp_oTextBox.Terminate()
        Tasks.Clear()
        Rows.Clear()
        Styles.Clear()
        Layers.Clear()
        Columns.Clear()
        TimeBlocks.Clear()
        Views.Clear()
    End Sub

    Public Sub CheckPredecessors()
        Dim i As Integer
        Dim oTask As clsTask
        For i = 1 To Tasks.Count
            oTask = Tasks.oCollection.m_oReturnArrayElement(i)
            oTask.mp_bWarning = False
        Next
        If Predecessors.Count = 0 Then
            Return
        End If
        Dim oPredecessor As clsPredecessor
        For i = 1 To Predecessors.Count
            oPredecessor = Predecessors.oCollection.m_oReturnArrayElement(i)
            oPredecessor.Check(mp_yPredecessorMode)
        Next
    End Sub

    Public Property EnforcePredecessors() As Boolean
        Get
            Return mp_bEnforcePredecessors
        End Get
        Set(ByVal value As Boolean)
            mp_bEnforcePredecessors = value
        End Set
    End Property

    Public Property PredecessorMode() As E_PREDECESSORMODE
        Get
            Return mp_yPredecessorMode
        End Get
        Set(ByVal value As E_PREDECESSORMODE)
            mp_yPredecessorMode = value
        End Set
    End Property

    Public Sub ForceEndTextEdit()
        mp_oTextBox.Terminate()
    End Sub

    Public ReadOnly Property ModuleCompletePath() As String
        Get
            Return System.Reflection.Assembly.GetExecutingAssembly.Location
        End Get
    End Property

    Public ReadOnly Property Version() As String
        Get
            Dim ai As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly
            Return ai.GetName.Version().ToString()
        End Get
    End Property

    Public Sub SaveToImage(ByVal Path As String)
        Dim fs As New System.IO.FileStream(Path, IO.FileMode.Create)
        Dim bmp As New RenderTargetBitmap(clsG.Width, clsG.Height, 96, 96, PixelFormats.Pbgra32)
        oCanvas.Measure(New System.Windows.Size(clsG.Width, clsG.Height))
        'oCanvas.Arrange(New Rect(0, 0, clsG.Width, clsG.Height))
        oCanvas.UpdateLayout()
        bmp.Render(oCanvas)
        Dim encoder As BitmapEncoder = New TiffBitmapEncoder()
        encoder.Frames.Add(BitmapFrame.Create(bmp))
        encoder.Save(fs)
        fs.Close()
        'Dim oBitmap As System.Drawing.Bitmap
        'oBitmap = New System.Drawing.Bitmap(clsG.Width, clsG.Height)
        'mp_oGraphics = System.Drawing.Graphics.FromImage(oBitmap)
        'mp_Draw()
        'oBitmap.Save(Path, Format)
    End Sub

    Public Sub AboutBox()
        Dim oForm As New fAbout()
        oForm.ShowDialog()
    End Sub

    Friend Sub mp_Redraw()
        Me.InvalidateVisual()
    End Sub

    Friend Sub mp_RedrawOtherThread()
        If Me.Dispatcher.HasShutdownFinished = False Then Return
        Dim oDraw As mp_RedrawHandler = AddressOf mp_Redraw
        Me.Dispatcher.Invoke(oDraw)
    End Sub

    Public Sub Redraw()
        Try
            mp_Redraw()
        Catch ex As Exception
            mp_RedrawOtherThread()
        End Try
    End Sub

    Public Property AllowSplitterMove() As Boolean
        Get
            Return mp_bAllowSplitterMove
        End Get
        Set(ByVal Value As Boolean)
            mp_bAllowSplitterMove = Value
        End Set
    End Property

    Public Property AllowPredecessorAdd() As Boolean
        Get
            Return mp_bAllowPredecessorAdd
        End Get
        Set(ByVal Value As Boolean)
            mp_bAllowPredecessorAdd = Value
        End Set
    End Property

    Public Property AllowAdd() As Boolean
        Get
            Return mp_bAllowAdd
        End Get
        Set(ByVal Value As Boolean)
            mp_bAllowAdd = Value
        End Set
    End Property

    Public Property AllowEdit() As Boolean
        Get
            Return mp_bAllowEdit
        End Get
        Set(ByVal Value As Boolean)
            mp_bAllowEdit = Value
        End Set
    End Property

    Public Property AllowColumnSize() As Boolean
        Get
            Return mp_bAllowColumnSize
        End Get
        Set(ByVal Value As Boolean)
            mp_bAllowColumnSize = Value
        End Set
    End Property

    Public Property AllowRowSize() As Boolean
        Get
            Return mp_bAllowRowSize
        End Get
        Set(ByVal Value As Boolean)
            mp_bAllowRowSize = Value
        End Set
    End Property

    Public Property AllowRowMove() As Boolean
        Get
            Return mp_bAllowRowMove
        End Get
        Set(ByVal Value As Boolean)
            mp_bAllowRowMove = Value
        End Set
    End Property

    Public Property AllowColumnMove() As Boolean
        Get
            Return mp_bAllowColumnMove
        End Get
        Set(ByVal Value As Boolean)
            mp_bAllowColumnMove = Value
        End Set
    End Property

    Public Property AllowTimeLineScroll() As Boolean
        Get
            Return mp_bAllowTimeLineScroll
        End Get
        Set(ByVal Value As Boolean)
            mp_bAllowTimeLineScroll = Value
        End Set
    End Property

    Public Property AddMode() As E_ADDMODE
        Get
            Return mp_yAddMode
        End Get
        Set(ByVal Value As E_ADDMODE)
            mp_yAddMode = Value
        End Set
    End Property

    Public Property AddDurationInterval() As E_INTERVAL
        Get
            Return mp_yAddDurationInterval
        End Get
        Set(ByVal value As E_INTERVAL)
            If Not (value = E_INTERVAL.IL_SECOND Or value = E_INTERVAL.IL_MINUTE Or value = E_INTERVAL.IL_HOUR Or value = E_INTERVAL.IL_DAY) Then
                mp_ErrorReport(SYS_ERRORS.INVALID_DURATION_INTERVAL, "Interval is invalid for a duration", "ActiveGanttVBWCtl.Set_AddDurationInterval")
                Return
            End If
            mp_yAddDurationInterval = value
        End Set
    End Property

    Public Property LayerEnableObjects() As E_LAYEROBJECTENABLE
        Get
            Return mp_yLayerEnableObjects
        End Get
        Set(ByVal Value As E_LAYEROBJECTENABLE)
            mp_yLayerEnableObjects = Value
        End Set
    End Property

    Public Property DoubleBuffering() As Boolean
        Get
            Return mp_bDoubleBuffering
        End Get
        Set(ByVal Value As Boolean)
            mp_bDoubleBuffering = Value
        End Set
    End Property

    Public Property MinRowHeight() As Integer
        Get
            Return mp_lMinRowHeight
        End Get
        Set(ByVal Value As Integer)
            mp_lMinRowHeight = Value
        End Set
    End Property

    Public Property MinColumnWidth() As Integer
        Get
            Return mp_lMinColumnWidth
        End Get
        Set(ByVal Value As Integer)
            mp_lMinColumnWidth = Value
        End Set
    End Property

    Public Property ControlTag() As String
        Get
            Return mp_sControlTag
        End Get
        Set(ByVal Value As String)
            mp_sControlTag = Value
        End Set
    End Property

    '//

    Private Sub mp_CHKPXPStart(ByVal clrBackground As Color)
        clsG.ClipRegion(0, 0, clsG.Width(), clsG.Height(), False)
        clsG.DrawLine(0, 0, clsG.Width(), clsG.Height(), GRE_LINETYPE.LT_FILLED, clrBackground, GRE_LINEDRAWSTYLE.LDS_SOLID)
    End Sub

    Private Sub mp_CHKPXPScrollButtons()
        mp_CHKPXPStart(Colors.Red)
        clsG.DrawScrollButton(20, 20, 17, 17, E_SCROLLBUTTON.SB_UP, E_SCROLLBUTTONSTATE.BS_NORMAL)
        clsG.DrawScrollButton(40, 20, 17, 17, E_SCROLLBUTTON.SB_UP, E_SCROLLBUTTONSTATE.BS_PUSHED)
        clsG.DrawScrollButton(60, 20, 17, 17, E_SCROLLBUTTON.SB_UP, E_SCROLLBUTTONSTATE.BS_INACTIVE)

        clsG.DrawScrollButton(20, 40, 17, 17, E_SCROLLBUTTON.SB_DOWN, E_SCROLLBUTTONSTATE.BS_NORMAL)
        clsG.DrawScrollButton(40, 40, 17, 17, E_SCROLLBUTTON.SB_DOWN, E_SCROLLBUTTONSTATE.BS_PUSHED)
        clsG.DrawScrollButton(60, 40, 17, 17, E_SCROLLBUTTON.SB_DOWN, E_SCROLLBUTTONSTATE.BS_INACTIVE)

        clsG.DrawScrollButton(20, 60, 17, 17, E_SCROLLBUTTON.SB_LEFT, E_SCROLLBUTTONSTATE.BS_NORMAL)
        clsG.DrawScrollButton(40, 60, 17, 17, E_SCROLLBUTTON.SB_LEFT, E_SCROLLBUTTONSTATE.BS_PUSHED)
        clsG.DrawScrollButton(60, 60, 17, 17, E_SCROLLBUTTON.SB_LEFT, E_SCROLLBUTTONSTATE.BS_INACTIVE)

        clsG.DrawScrollButton(20, 80, 17, 17, E_SCROLLBUTTON.SB_RIGHT, E_SCROLLBUTTONSTATE.BS_NORMAL)
        clsG.DrawScrollButton(40, 80, 17, 17, E_SCROLLBUTTON.SB_RIGHT, E_SCROLLBUTTONSTATE.BS_PUSHED)
        clsG.DrawScrollButton(60, 80, 17, 17, E_SCROLLBUTTON.SB_RIGHT, E_SCROLLBUTTONSTATE.BS_INACTIVE)
    End Sub

    Private Sub mp_CHKPXPLines()
        mp_CHKPXPStart(Colors.White)
        clsG.DrawLine(20, 10, 60, 10, GRE_LINETYPE.LT_NORMAL, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)
        clsG.DrawLine(30, 20, 30, 70, GRE_LINETYPE.LT_NORMAL, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)
        clsG.DrawLine(40, 40, 60, 60, GRE_LINETYPE.LT_BORDER, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)
        clsG.DrawLine(70, 70, 90, 90, GRE_LINETYPE.LT_FILLED, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)
    End Sub

    Private Sub mp_CHKPXPButtons()
        mp_CHKPXPStart(Colors.Red)
        clsG.DrawButton(New Rect(20, 20, 100, 50), E_SCROLLBUTTONSTATE.BS_NORMAL)
        clsG.mp_DrawItem(50, 50, 100, 100, "", "", False, Nothing, 100, 100, Styles.FItem("DS_TICKMARKAREA"))
    End Sub

    Private Sub mp_CHKPXPArrows()
        mp_CHKPXPStart(Colors.White)
        clsG.mp_DrawArrow(100, 120, GRE_ARROWDIRECTION.AWD_RIGHT, 20, Colors.Black)
        clsG.mp_DrawArrow(50, 120, GRE_ARROWDIRECTION.AWD_LEFT, 20, Colors.Black)
        clsG.mp_DrawArrow(100, 150, GRE_ARROWDIRECTION.AWD_UP, 20, Colors.Black)
        clsG.mp_DrawArrow(50, 170, GRE_ARROWDIRECTION.AWD_DOWN, 20, Colors.Black)
    End Sub

    Private Sub mp_CHKPXPFigures()
        mp_CHKPXPStart(Colors.White)
        clsG.DrawFigure(200, 20, 20, 20, GRE_FIGURETYPE.FT_ARROWDOWN, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)
        clsG.DrawFigure(230, 20, 20, 20, GRE_FIGURETYPE.FT_ARROWUP, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)

        clsG.DrawFigure(200, 60, 20, 20, GRE_FIGURETYPE.FT_CIRCLEARROWDOWN, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)
        clsG.DrawFigure(230, 60, 20, 20, GRE_FIGURETYPE.FT_CIRCLEARROWUP, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)

        clsG.DrawFigure(200, 100, 20, 20, GRE_FIGURETYPE.FT_CIRCLETRIANGLEDOWN, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)
        clsG.DrawFigure(230, 100, 20, 20, GRE_FIGURETYPE.FT_CIRCLETRIANGLEUP, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)

        clsG.DrawFigure(200, 140, 20, 20, GRE_FIGURETYPE.FT_PROJECTDOWN, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)
        clsG.DrawFigure(230, 140, 20, 20, GRE_FIGURETYPE.FT_PROJECTUP, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)

        clsG.DrawFigure(200, 180, 20, 20, GRE_FIGURETYPE.FT_SMALLPROJECTDOWN, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)
        clsG.DrawFigure(230, 180, 20, 20, GRE_FIGURETYPE.FT_SMALLPROJECTUP, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)

        clsG.DrawFigure(200, 220, 20, 20, GRE_FIGURETYPE.FT_TRIANGLEDOWN, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)
        clsG.DrawFigure(230, 220, 20, 20, GRE_FIGURETYPE.FT_TRIANGLEUP, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)

        clsG.DrawFigure(200, 260, 20, 20, GRE_FIGURETYPE.FT_TRIANGLELEFT, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)
        clsG.DrawFigure(230, 260, 20, 20, GRE_FIGURETYPE.FT_TRIANGLERIGHT, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)

        clsG.DrawFigure(200, 300, 20, 20, GRE_FIGURETYPE.FT_SQUARE, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)
        clsG.DrawFigure(230, 300, 20, 20, GRE_FIGURETYPE.FT_CIRCLE, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)

        clsG.DrawFigure(200, 340, 20, 20, GRE_FIGURETYPE.FT_DIAMOND, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)
        clsG.DrawFigure(230, 340, 20, 20, GRE_FIGURETYPE.FT_RECTANGLE, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)

        clsG.DrawFigure(200, 380, 20, 20, GRE_FIGURETYPE.FT_NONE, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)
        clsG.DrawFigure(230, 380, 20, 20, GRE_FIGURETYPE.FT_NONE, Colors.Blue, Colors.Black, GRE_LINEDRAWSTYLE.LDS_SOLID)
    End Sub

    Private Sub mp_CHKPXPGradients()
        mp_CHKPXPStart(Colors.White)
        clsG.GradientFill(20, 250, 120, 320, Colors.Black, Colors.Blue, GRE_GRADIENTFILLMODE.GDT_HORIZONTAL)
        clsG.GradientFill(20, 340, 120, 400, Colors.Black, Colors.Blue, GRE_GRADIENTFILLMODE.GDT_VERTICAL)
    End Sub

    Private Sub mp_CHKPXPText()
        mp_CHKPXPStart(Colors.White)
        Dim oFlags As New clsTextFlags(Me)
        oFlags.HorizontalAlignment = GRE_HORIZONTALALIGNMENT.HAL_CENTER
        oFlags.VerticalAlignment = GRE_VERTICALALIGNMENT.VAL_CENTER
        'clsG.DrawTextEx(300, 100, 400, 400, "M Hello World", oFlags, Colors.Black, New Font("Arial", 10), True)
        clsG.DrawTextEx(200, 100, 400, 300, "M Hello World", oFlags, Colors.Black, New Font("Arial", 10), True)
    End Sub

    Private Sub mp_CHKPXPHatch()
        mp_CHKPXPStart(Colors.White)
        Dim i As Integer
        Dim j As Integer
        j = 0
        For i = 0 To 9
            clsG.HatchFill(0, j * 30, 100, (j * 30) + 20, Colors.Black, Colors.White, i)
            j = j + 1
        Next
        j = 0
        For i = 10 To 19
            clsG.HatchFill(120, j * 30, 220, (j * 30) + 20, Colors.Black, Colors.White, i)
            j = j + 1
        Next

        j = 0
        For i = 20 To 29
            clsG.HatchFill(240, j * 30, 340, (j * 30) + 20, Colors.Black, Colors.White, i)
            j = j + 1
        Next
        j = 0
        For i = 30 To 39
            clsG.HatchFill(360, j * 30, 460, (j * 30) + 20, Colors.Black, Colors.White, i)
            j = j + 1
        Next
        j = 0
        For i = 40 To 49
            clsG.HatchFill(480, j * 30, 580, (j * 30) + 20, Colors.Black, Colors.White, i)
            j = j + 1
        Next
        j = 0
        For i = 50 To 52
            clsG.HatchFill(600, j * 30, 700, (j * 30) + 20, Colors.Black, Colors.White, i)
            j = j + 1
        Next
    End Sub

    Friend Function ShowAbout() As Boolean
        If System.ComponentModel.DesignerProperties.GetIsInDesignMode(Me) = False Then
            Return False
        End If
        Dim oKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("Licenses\BF5980F3-B7A2-4fcf-9E12-575D378FDFA7")
        If Not oKey Is Nothing Then
            Dim sDefault As String = CType(oKey.GetValue(""), String)
            If Not sDefault Is Nothing Then
                If String.Compare("15EFB0FA-E747-46ed-993F-D76EFF641B02", sDefault, False) = 0 Then
                    Return False
                End If
            End If
        End If
        Return True
    End Function

End Class

Public Class AGError
    Inherits Exception

    Private mp_sErrDescription As String
    Private mp_lErrNumber As Integer
    Private mp_sErrSource As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal s As String)
        MyBase.New(s)
    End Sub

    Public Sub New(ByVal s As String, ByVal ex As Exception)
        MyBase.New(s, ex)
    End Sub

    Public Property ErrDescription() As String
        Get
            Return mp_sErrDescription
        End Get
        Set(ByVal value As String)
            mp_sErrDescription = value
        End Set
    End Property

    Public Property ErrNumber() As Integer
        Get
            Return mp_lErrNumber
        End Get
        Set(ByVal value As Integer)
            mp_lErrNumber = value
        End Set
    End Property

    Public Property ErrSource() As String
        Get
            Return mp_sErrSource
        End Get
        Set(ByVal value As String)
            mp_sErrSource = value
        End Set
    End Property
End Class
