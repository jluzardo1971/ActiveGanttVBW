Public Class clsPrinter

    Private mp_oControl As ActiveGanttVBWCtl
    Private mp_dtPrintAreaStartDate As AGVBW.DateTime
    Private mp_dtPrintAreaEndDate As AGVBW.DateTime
    Private mp_dtPrintStartDateBuffer As AGVBW.DateTime
    Private mp_oView As clsView

    Friend Sub New(ByVal Value As ActiveGanttVBWCtl)
        mp_oControl = Value
    End Sub

    Public Sub Initialize(ByVal StartDate As AGVBW.DateTime, ByVal EndDate As AGVBW.DateTime, Optional ByVal ControlHeight As Integer = -1)
        Const CorrectionFactor As Integer = 5

        mp_oView = mp_oControl.CurrentViewObject
        mp_dtPrintAreaStartDate = StartDate
        mp_dtPrintAreaEndDate = EndDate
        mp_oControl.clsG.CustomWidth = mp_oControl.MathLib.DateTimeDiff(mp_oView.Interval, StartDate, EndDate) / mp_oView.Factor + CorrectionFactor
        mp_oControl.clsG.CustomWidth = mp_oControl.clsG.CustomWidth + mp_oControl.Splitter.Right
        If ControlHeight = -1 Then
            mp_oControl.clsG.CustomHeight = mp_oControl.Rows.Height + (mp_oControl.Rows.Count * 1) + mp_oControl.CurrentViewObject.ClientArea.Top + mp_oControl.mt_BorderThickness
            If mp_oControl.clsG.CustomHeight < mp_oControl.f_Height Then
                mp_oControl.clsG.CustomHeight = mp_oControl.f_Height
            End If
        Else
            mp_oControl.clsG.CustomHeight = mp_oControl.f_Height
        End If
        If mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.Enabled = False Then
            mp_dtPrintStartDateBuffer = mp_oView.TimeLine.StartDate
            mp_oView.TimeLine.f_StartDate = mp_dtPrintAreaStartDate
        Else
            mp_dtPrintStartDateBuffer = mp_oView.TimeLine.TimeLineScrollBar.StartDate
            mp_oView.TimeLine.TimeLineScrollBar.StartDate = mp_dtPrintAreaStartDate
        End If
    End Sub

    Public Sub Terminate()
        If mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.Enabled = False Then
            mp_oView.TimeLine.f_StartDate = mp_dtPrintStartDateBuffer
        Else
            mp_oView.TimeLine.TimeLineScrollBar.StartDate = mp_dtPrintStartDateBuffer
        End If
    End Sub

    Public Sub PrintControl(ByRef r_Visual As Visual, ByVal XOrigin As Integer, ByVal YOrigin As Integer, ByVal XOriginExtents As Integer, ByVal YOriginExtents As Integer, ByVal MarginX As Integer, ByVal MarginY As Integer, ByVal DestScale As Integer)
        mp_oControl.clsG.CustomPrinting = True
        mp_oControl.mp_PositionScrollBars()
        Dim oVisual As New DrawingVisual
        Dim oTransformMargins As New TranslateTransform()
        Dim oScaleTransform As New ScaleTransform
        Dim oTransformGroup As New TransformGroup
        Dim oClip As New RectangleGeometry(New Rect(XOrigin, YOrigin, XOriginExtents, YOriginExtents))
        oTransformMargins.X = MarginX - XOrigin
        oTransformMargins.Y = MarginY - YOrigin
        mp_oControl.clsG.CustomDC = oVisual.RenderOpen()
        mp_oControl.f_Draw()
        mp_oControl.clsG.CustomDC.Close()
        mp_oControl.clsG.CustomPrinting = False
        mp_oControl.mp_PositionScrollBars()
        oVisual.Clip = oClip
        oVisual.Transform = oTransformMargins
        r_Visual = oVisual
    End Sub

    Public Function GetVisual() As Visual
        mp_oControl.clsG.CustomPrinting = True
        mp_oControl.mp_PositionScrollBars()
        Dim oVisual As New DrawingVisual
        mp_oControl.clsG.CustomDC = oVisual.RenderOpen()
        mp_oControl.f_Draw()
        mp_oControl.clsG.CustomDC.Close()
        mp_oControl.clsG.CustomPrinting = False
        mp_oControl.mp_PositionScrollBars()
        Return oVisual
    End Function

    Public ReadOnly Property PrintAreaEndDate() As AGVBW.DateTime
        Get
            Return mp_dtPrintAreaEndDate
        End Get
    End Property

    Public ReadOnly Property PrintAreaStartDate() As AGVBW.DateTime
        Get
            Return mp_dtPrintAreaStartDate
        End Get
    End Property

    Public ReadOnly Property PrintAreaWidth() As Integer
        Get
            Return mp_oControl.clsG.CustomWidth
        End Get
    End Property

    Public ReadOnly Property PrintAreaHeight() As Integer
        Get
            Return mp_oControl.clsG.CustomHeight
        End Get
    End Property

End Class
