Public Class clsToolTip

    Private mp_oControl As ActiveGanttVBWCtl
    Private mp_lLeft As Integer
    Private mp_lTop As Integer
    Private mp_lWidth As Integer
    Private mp_lHeight As Integer
    Private mp_sText As String
    Private mp_bVisible As Boolean
    Private mp_bBackupDCActive As Boolean
    Private mp_oFont As Font
    Private mp_lBackupLeft As Integer
    Private mp_lBackupTop As Integer
    Private mp_lBackupRight As Integer
    Private mp_lBackupBottom As Integer
    Private mp_bAutomaticSizing As Boolean = False
    Private mp_clrBackColor As System.Windows.Media.Color = System.Windows.Media.Colors.LightYellow
    Private mp_clrForeColor As System.Windows.Media.Color = System.Windows.Media.Colors.Black
    Private mp_clrBorderColor As System.Windows.Media.Color = System.Windows.Media.Colors.Black
    Private mp_oToolTipCanvas As Canvas
    Private mp_lToolTipCanvasIndex As Integer = -1

    Friend Sub New(ByVal Value As ActiveGanttVBWCtl)
        mp_oControl = Value
        mp_oFont = New Font("Tahoma", 8)
        mp_oToolTipCanvas = New Canvas()
    End Sub

    Public Property Font() As Font
        Get
            Return mp_oFont
        End Get
        Set(ByVal Value As Font)
            mp_oFont = Value
        End Set
    End Property

    Public Property BackColor() As System.Windows.Media.Color
        Get
            Return mp_clrBackColor
        End Get
        Set(ByVal Value As System.Windows.Media.Color)
            mp_clrBackColor = Value
        End Set
    End Property

    Public Property ForeColor() As System.Windows.Media.Color
        Get
            Return mp_clrForeColor
        End Get
        Set(ByVal Value As System.Windows.Media.Color)
            mp_clrForeColor = Value
        End Set
    End Property

    Public Property BorderColor() As System.Windows.Media.Color
        Get
            Return mp_clrBorderColor
        End Get
        Set(ByVal Value As System.Windows.Media.Color)
            mp_clrBorderColor = Value
        End Set
    End Property

    Public Property Text() As String
        Get
            Return mp_sText
        End Get
        Set(ByVal Value As String)
            mp_sText = Value
            If mp_bAutomaticSizing = True Then
                Dim oTypeFace As New Typeface(Font.FamilyName)
                Dim oFormattedText As New FormattedText(mp_sText, mp_oControl.Culture, FlowDirection.LeftToRight, oTypeFace, Font.WPFFontSize, New SolidColorBrush(Colors.Black))
                mp_lWidth = oFormattedText.Width
                mp_lHeight = oFormattedText.Height
            End If
        End Set
    End Property

    Public Property AutomaticSizing() As Boolean
        Get
            Return mp_bAutomaticSizing
        End Get
        Set(ByVal Value As Boolean)
            mp_bAutomaticSizing = Value
        End Set
    End Property

    Public Property Left() As Integer
        Get
            Return mp_lLeft
        End Get
        Set(ByVal Value As Integer)
            mp_lLeft = Value
        End Set
    End Property

    Public ReadOnly Property Right() As Integer
        Get
            Return mp_lLeft + mp_lWidth
        End Get
    End Property

    Public Property Top() As Integer
        Get
            Return mp_lTop
        End Get
        Set(ByVal Value As Integer)
            mp_lTop = Value
        End Set
    End Property

    Public ReadOnly Property Bottom() As Integer
        Get
            Return mp_lTop + mp_lHeight
        End Get
    End Property

    Public Property Width() As Integer
        Get
            Return mp_lWidth
        End Get
        Set(ByVal Value As Integer)
            mp_lWidth = Value
        End Set
    End Property

    Public Property Height() As Integer
        Get
            Return mp_lHeight
        End Get
        Set(ByVal Value As Integer)
            mp_lHeight = Value
        End Set
    End Property

    Public Property Visible() As Boolean
        Get
            Return mp_bVisible
        End Get
        Set(ByVal Value As Boolean)

            mp_bVisible = Value
            If (mp_lWidth = 0 Or mp_lHeight = 0) Then
                mp_bVisible = False
            End If
            If (mp_bVisible = True) Then
                Dim oRectangle As New Rectangle
                If mp_oControl.f_Canvas.Children.Count = 0 Then

                    oRectangle = New Rectangle
                    oRectangle.SetValue(Canvas.LeftProperty, CDbl(0))
                    oRectangle.SetValue(Canvas.TopProperty, CDbl(0))
                    oRectangle.Width = mp_lWidth
                    oRectangle.Height = mp_lHeight
                    oRectangle.Fill = mp_oControl.clsG.GetBrush(mp_clrBackColor)
                    oRectangle.Stroke = mp_oControl.clsG.GetBrush(mp_clrBorderColor)
                    oRectangle.StrokeThickness = 1
                    mp_oToolTipCanvas.Children.Add(oRectangle)
                    mp_oToolTipCanvas.SetValue(Canvas.LeftProperty, CDbl(mp_lLeft))
                    mp_oToolTipCanvas.SetValue(Canvas.TopProperty, CDbl(mp_lTop))
                    mp_oToolTipCanvas.Width = mp_lWidth
                    mp_oToolTipCanvas.Height = mp_lHeight
                    mp_lToolTipCanvasIndex = mp_oControl.f_Canvas.Children.Add(mp_oToolTipCanvas)
                Else
                    mp_oToolTipCanvas.Children.Clear()
                    oRectangle = New Rectangle
                    oRectangle.SetValue(Canvas.LeftProperty, CDbl(0))
                    oRectangle.SetValue(Canvas.TopProperty, CDbl(0))
                    oRectangle.Width = mp_lWidth
                    oRectangle.Height = mp_lHeight
                    oRectangle.Fill = mp_oControl.clsG.GetBrush(mp_clrBackColor)
                    oRectangle.Stroke = mp_oControl.clsG.GetBrush(mp_clrBorderColor)
                    oRectangle.StrokeThickness = 1
                    mp_oToolTipCanvas.Children.Add(oRectangle)
                    mp_oToolTipCanvas.Width = mp_lWidth
                    mp_oToolTipCanvas.Height = mp_lHeight
                End If
                Select Case mp_oControl.ToolTipEventArgs.ToolTipType
                    Case E_TOOLTIPTYPE.TPT_HOVER
                        mp_oControl.ToolTipEventArgs.Graphics = mp_oToolTipCanvas
                        mp_oControl.ToolTipEventArgs.CustomDraw = False
                        mp_oControl.FireOnMouseHoverToolTipDraw(mp_oControl.ToolTipEventArgs.EventTarget)
                    Case E_TOOLTIPTYPE.TPT_MOVEMENT
                        mp_oControl.ToolTipEventArgs.Graphics = mp_oToolTipCanvas
                        mp_oControl.ToolTipEventArgs.CustomDraw = False
                        mp_oControl.FireOnMouseMoveToolTipDraw(mp_oControl.ToolTipEventArgs.Operation)
                End Select
                If mp_oControl.ToolTipEventArgs.CustomDraw = False Then
                    'lControlGraphics.DrawString(mp_sText, mp_oFont, New SolidBrush(mp_clrForeColor), mp_lLeft, mp_lTop)
                End If
            Else
                If mp_lToolTipCanvasIndex > -1 Then
                    mp_oControl.f_Canvas.Children.Remove(mp_oToolTipCanvas)
                    mp_lToolTipCanvasIndex = -1
                End If
            End If
        End Set
    End Property

    Private Sub RestoreBackupDC()

    End Sub

End Class
