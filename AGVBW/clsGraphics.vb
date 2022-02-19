Option Explicit On

Friend Class clsGraphics

    Private Structure T_PRECT
        Public lLeft As Integer
        Public lTop As Integer
        Public lRight As Integer
        Public lBottom As Integer
    End Structure

    Private Enum T_HATCHTYPE
        HT_RECTANGLE = 1
        HT_LINE = 2
    End Enum

    Private mp_oControl As ActiveGanttVBWCtl
    Private mp_udtPreviousClipRegion As T_PRECT
    Private mp_audtActiveReversibleFrames As System.Collections.ArrayList
    Private mp_audtActiveReversibleLinesStart As System.Collections.ArrayList
    Private mp_audtActiveReversibleLinesEnd As System.Collections.ArrayList
    Private mp_bCustomPrinting As Boolean
    Private mp_lCustomDC As DrawingContext
    Private mp_lPWidth As Integer
    Private mp_lPHeight As Integer
    Private mp_lFocusLeft As Integer
    Private mp_lFocusTop As Integer
    Private mp_lFocusRight As Integer
    Private mp_lFocusBottom As Integer
    Private mp_bEnableClipRegions As Boolean
    Friend mp_oToolTipGraphics As DrawingContext
    Friend bToolTipGraphics As Boolean
    Private mp_bRequiresPop As Boolean = False

    Private mp_oSelectionLine As Line
    Private mp_oSelectionRectangle As Rectangle

    Private mp_lSelectionRectangleIndex As Integer = -1
    Private mp_lSelectionLineIndex As Integer = -1

    Friend mp_oTextFinalLayout As Rect

    '// ---------------------------------------------------------------------------------------------------------------------
    '// Construction/Destruction & Initialization
    '// ---------------------------------------------------------------------------------------------------------------------

    Friend Sub New(ByVal Value As ActiveGanttVBWCtl)
        mp_oControl = Value
        mp_audtActiveReversibleFrames = New System.Collections.ArrayList()
        mp_audtActiveReversibleLinesStart = New System.Collections.ArrayList()
        mp_audtActiveReversibleLinesEnd = New System.Collections.ArrayList()
        mp_bCustomPrinting = False
        mp_bEnableClipRegions = True
        bToolTipGraphics = False
        mp_oSelectionLine = New Line()
        mp_oSelectionRectangle = New Rectangle()
    End Sub

    Friend Property RequiresPop() As Boolean
        Get
            Return mp_bRequiresPop
        End Get
        Set(value As Boolean)
            mp_bRequiresPop = value
        End Set
    End Property

    Friend Property EnableClipRegions() As Boolean
        Get
            Return mp_bEnableClipRegions
        End Get
        Set(ByVal Value As Boolean)
            mp_bEnableClipRegions = Value
        End Set
    End Property

    Friend Property f_FocusLeft() As Integer
        Get
            Return mp_lFocusLeft
        End Get
        Set(ByVal Value As Integer)
            mp_lFocusLeft = Value
        End Set
    End Property

    Friend Property f_FocusTop() As Integer
        Get
            Return mp_lFocusTop
        End Get
        Set(ByVal Value As Integer)
            mp_lFocusTop = Value
        End Set
    End Property

    Friend Property f_FocusRight() As Integer
        Get
            Return mp_lFocusRight
        End Get
        Set(ByVal Value As Integer)
            mp_lFocusRight = Value
        End Set
    End Property

    Friend Property f_FocusBottom() As Integer
        Get
            Return mp_lFocusBottom
        End Get
        Set(ByVal Value As Integer)
            mp_lFocusBottom = Value
        End Set
    End Property

    Public ReadOnly Property oGraphics() As DrawingContext
        Get
            If mp_bCustomPrinting = False Then
                If bToolTipGraphics = False Then
                    Return mp_oControl.f_HDC()
                Else
                    Return mp_oToolTipGraphics
                End If
            Else
                Return mp_lCustomDC
            End If
        End Get
    End Property

    Public Property CustomPrinting() As Boolean
        Get
            Return mp_bCustomPrinting
        End Get
        Set(ByVal Value As Boolean)
            mp_bCustomPrinting = Value
        End Set
    End Property

    Public Property CustomDC() As DrawingContext
        Set(ByVal Value As DrawingContext)
            mp_lCustomDC = Value
        End Set
        Get
            Return mp_lCustomDC
        End Get
    End Property

    Public Function Width() As Integer
        If mp_bCustomPrinting = False Then
            Return mp_oControl.f_Width
        Else
            Return mp_lPWidth
        End If
    End Function

    Public Function Height() As Integer
        If mp_bCustomPrinting = False Then
            Return mp_oControl.f_Height
        Else
            Return mp_lPHeight
        End If
    End Function

    Public Property CustomWidth() As Integer
        Get
            Return mp_lPWidth
        End Get
        Set(ByVal Value As Integer)
            mp_lPWidth = Value
        End Set
    End Property

    Public Property CustomHeight() As Integer
        Get
            Return mp_lPHeight
        End Get
        Set(ByVal Value As Integer)
            mp_lPHeight = Value
        End Set
    End Property

    Public Sub DrawPolygon(ByVal v_lColor As Color, ByRef r_oPoints() As Point, Optional ByVal bFilled As Boolean = False)
        Dim oPathFigure As New PathFigure()
        Dim oPathSegmentCollection As New PathSegmentCollection()
        Dim oPathGeometry As New PathGeometry
        Dim i As Integer
        oPathFigure.StartPoint = r_oPoints(0)
        For i = 0 To r_oPoints.GetUpperBound(0)
            oPathSegmentCollection.Add(New LineSegment(r_oPoints(i), False))
        Next
        oPathSegmentCollection.Add(New LineSegment(r_oPoints(0), False))
        oPathFigure.Segments = oPathSegmentCollection
        oPathGeometry.Figures.Add(oPathFigure)

        oPathGeometry.Freeze()
        If bFilled = False Then
            oGraphics.DrawGeometry(Nothing, GetPen(v_lColor), oPathGeometry)
        Else
            oGraphics.DrawGeometry(GetBrush(v_lColor), Nothing, oPathGeometry)
        End If
    End Sub

    Public Function GetPen(ByVal oColor As Color) As Pen
        Dim oBrush As New SolidColorBrush()
        Dim oPen As New Pen()
        oBrush.Color = oColor
        oBrush.Freeze()
        oPen.Brush = oBrush
        oPen.Thickness = 1
        oPen.Freeze()
        Return oPen
    End Function

    Public Function GetBrush(ByVal oColor As Color) As SolidColorBrush
        Dim oBrush As New SolidColorBrush(oColor)
        oBrush.Freeze()
        Return oBrush
    End Function

    Public Function ConvertColor(ByVal dwOleColour As Color) As Integer
        '        Dim clrref As Integer
        '        OleTranslateColor(dwOleColour, 0, clrref)
        '        ConvertColor = clrref
    End Function

    Public Sub DrawEdge(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal clrBackColor As Color, ByVal v_yButtonStyle As GRE_BUTTONSTYLE, ByVal v_lEdgeType As GRE_EDGETYPE, ByVal v_bFilled As Boolean, ByVal oStyle As clsStyle)
        Dim lExteriorLeftTopColor As Color
        Dim lInteriorLeftTopColor As Color
        Dim lExteriorRightBottomColor As Color
        Dim lInteriorRightBottomColor As Color
        If v_yButtonStyle = GRE_BUTTONSTYLE.BT_NORMALWINDOWS Then
            Select Case v_lEdgeType
                Case GRE_EDGETYPE.ET_RAISED
                    If oStyle Is Nothing Then
                        lExteriorLeftTopColor = Color.FromArgb(255, 240, 240, 240)
                        lInteriorLeftTopColor = Color.FromArgb(255, 192, 192, 192)
                        lInteriorRightBottomColor = Colors.Gray
                        lExteriorRightBottomColor = Color.FromArgb(255, 64, 64, 64)
                    Else
                        lExteriorLeftTopColor = oStyle.ButtonBorderStyle.RaisedExteriorLeftTopColor
                        lInteriorLeftTopColor = oStyle.ButtonBorderStyle.RaisedInteriorLeftTopColor
                        lInteriorRightBottomColor = oStyle.ButtonBorderStyle.RaisedInteriorRightBottomColor
                        lExteriorRightBottomColor = oStyle.ButtonBorderStyle.RaisedExteriorRightBottomColor
                    End If
                Case GRE_EDGETYPE.ET_SUNKEN
                    If oStyle Is Nothing Then
                        lExteriorLeftTopColor = Colors.Gray
                        lInteriorLeftTopColor = Color.FromArgb(255, 64, 64, 64)
                        lInteriorRightBottomColor = Color.FromArgb(255, 192, 192, 192)
                        lExteriorRightBottomColor = Color.FromArgb(255, 240, 240, 240)
                    Else
                        lExteriorLeftTopColor = oStyle.ButtonBorderStyle.SunkenExteriorLeftTopColor
                        lInteriorLeftTopColor = oStyle.ButtonBorderStyle.SunkenInteriorLeftTopColor
                        lInteriorRightBottomColor = oStyle.ButtonBorderStyle.SunkenInteriorRightBottomColor
                        lExteriorRightBottomColor = oStyle.ButtonBorderStyle.SunkenExteriorRightBottomColor
                    End If
            End Select
            '// Exterior Left
            DrawLine(v_X1, v_Y1, v_X1, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Exterior Top
            DrawLine(v_X1, v_Y1, v_X2, v_Y1, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Exterior Right
            DrawLine(v_X2, v_Y2, v_X2, v_Y1, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Exterior Bottom
            DrawLine(v_X1, v_Y2, v_X2, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Interior Left
            DrawLine(v_X1 + 1, v_Y1 + 1, v_X1 + 1, v_Y2 - 1, GRE_LINETYPE.LT_NORMAL, lInteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Interior Top
            DrawLine(v_X1 + 1, v_Y1 + 1, v_X2 - 1, v_Y1 + 1, GRE_LINETYPE.LT_NORMAL, lInteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Interior Right
            DrawLine(v_X2 - 1, v_Y2 - 1, v_X2 - 1, v_Y1 + 1, GRE_LINETYPE.LT_NORMAL, lInteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Interior Bottom
            DrawLine(v_X1 + 1, v_Y2 - 1, v_X2 - 1, v_Y2 - 1, GRE_LINETYPE.LT_NORMAL, lInteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            If v_bFilled = True Then
                DrawLine(v_X1 + 2, v_Y1 + 2, v_X2 - 2, v_Y2 - 2, GRE_LINETYPE.LT_FILLED, clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            End If
        Else
            Select Case v_lEdgeType
                Case GRE_EDGETYPE.ET_RAISED
                    If oStyle Is Nothing Then
                        lExteriorLeftTopColor = Colors.White
                        lExteriorRightBottomColor = Color.FromArgb(255, 64, 64, 64)
                    Else
                        lExteriorLeftTopColor = oStyle.ButtonBorderStyle.RaisedExteriorLeftTopColor
                        lExteriorRightBottomColor = oStyle.ButtonBorderStyle.RaisedExteriorRightBottomColor
                    End If
                Case GRE_EDGETYPE.ET_SUNKEN
                    If oStyle Is Nothing Then
                        lExteriorLeftTopColor = Colors.Gray
                        lExteriorRightBottomColor = Colors.WhiteSmoke
                    Else
                        lExteriorLeftTopColor = oStyle.ButtonBorderStyle.SunkenExteriorLeftTopColor
                        lExteriorRightBottomColor = oStyle.ButtonBorderStyle.SunkenExteriorRightBottomColor
                    End If
            End Select
            DrawLine(v_X1, v_Y1, v_X2, v_Y1, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            DrawLine(v_X1, v_Y1, v_X1, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            DrawLine(v_X1, v_Y2, v_X2, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            DrawLine(v_X2, v_Y2, v_X2, v_Y1 - 1, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            If v_bFilled = True Then
                DrawLine(v_X1 + 1, v_Y1 + 1, v_X2 - 1, v_Y2 - 1, GRE_LINETYPE.LT_FILLED, clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            End If
        End If
    End Sub

    Public Sub DrawLine(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_yStyle As GRE_LINETYPE, ByVal v_lColor As Color, ByVal v_lDrawStyle As GRE_LINEDRAWSTYLE)
        DrawLine(v_X1, v_Y1, v_X2, v_Y2, v_yStyle, v_lColor, v_lDrawStyle, 1, True)
    End Sub

    Public Sub DrawLine(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_yStyle As GRE_LINETYPE, ByVal v_lColor As Color, ByVal v_lDrawStyle As GRE_LINEDRAWSTYLE, ByVal v_lWidth As Integer)
        DrawLine(v_X1, v_Y1, v_X2, v_Y2, v_yStyle, v_lColor, v_lDrawStyle, v_lWidth, True)
    End Sub

    Public Sub CorrectRectCoords(ByRef X1 As Integer, ByRef Y1 As Integer, ByRef X2 As Integer, ByRef Y2 As Integer)
        Dim iBuff As Integer = 0
        If (X2 - X1) < 0 Then
            iBuff = X1
            X1 = X2
            X2 = iBuff
        End If
        If (Y2 - Y1) < 0 Then
            iBuff = Y1
            Y1 = Y2
            Y2 = iBuff
        End If
    End Sub

    Public Sub DrawLine(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_yStyle As GRE_LINETYPE, ByVal v_lColor As Color, ByVal v_lDrawStyle As GRE_LINEDRAWSTYLE, ByVal v_lWidth As Integer, ByVal v_bCreatePens As Boolean)
        CorrectRectCoords(v_X1, v_Y1, v_X2, v_Y2)
        Select Case v_yStyle
            Case GRE_LINETYPE.LT_NORMAL
                If v_X1 <> v_X2 Then
                    v_X2 = v_X2 + 1
                End If
                If v_Y1 <> v_Y2 Then
                    v_Y2 = v_Y2 + 1
                End If
                If v_X1 = v_X2 Then
                    v_X1 = v_X1 + 1
                    v_X2 = v_X2 + 1
                End If
                If v_Y1 = v_Y2 Then
                    v_Y1 = v_Y1 + 1
                    v_Y2 = v_Y2 + 1
                End If
                If v_lDrawStyle = GRE_LINEDRAWSTYLE.LDS_SOLID Then
                    oGraphics.DrawLine(GetPen(v_lColor), New Point(v_X1, v_Y1), New Point(v_X2, v_Y2))
                ElseIf v_lDrawStyle = GRE_LINEDRAWSTYLE.LDS_DOT Then
                    Dim oPen As New Pen(GetBrush(Colors.Gray), 1)
                    oPen.DashStyle = DashStyles.Dot
                    oGraphics.DrawLine(oPen, New Point(v_X1, v_Y1), New Point(v_X2, v_Y2))
                End If
            Case GRE_LINETYPE.LT_BORDER
                If v_lDrawStyle = GRE_LINEDRAWSTYLE.LDS_SOLID Then
                    oGraphics.DrawRectangle(Nothing, GetPen(v_lColor), New Rect(New Point(v_X1 + 1, v_Y1 + 1), New Point(v_X2 + 1, v_Y2 + 1)))
                ElseIf v_lDrawStyle = GRE_LINEDRAWSTYLE.LDS_DOT Then
                    Dim oPen As New Pen(GetBrush(Colors.Gray), 1)
                    oPen.DashStyle = DashStyles.Dot
                    oGraphics.DrawRectangle(Nothing, oPen, New Rect(New Point(v_X1 + 1, v_Y1 + 1), New Point(v_X2 + 1, v_Y2 + 1)))
                End If
            Case GRE_LINETYPE.LT_FILLED
                If v_X1 <> v_X2 Then
                    v_X2 = v_X2 - 1
                End If
                If v_Y1 <> v_Y2 Then
                    v_Y2 = v_Y2 - 1
                End If
                oGraphics.DrawRectangle(GetBrush(v_lColor), Nothing, New Rect(v_X1, v_Y1, v_X2 - v_X1 + 2, v_Y2 - v_Y1 + 2))
        End Select
    End Sub

    Public Sub DrawFigure(ByVal v_X As Integer, ByVal v_Y As Integer, ByVal v_dx As Integer, ByVal v_dy As Integer, ByVal v_yFigureType As GRE_FIGURETYPE, ByVal v_lBorderColor As Color, ByVal v_lFillColor As Color, ByVal v_yBorderStyle As GRE_LINEDRAWSTYLE)
        If v_dx Mod 2 <> 0 Then
            v_dx = v_dx + 1
            v_dy = v_dy + 1
        End If
        Select Case v_yFigureType
            Case GRE_FIGURETYPE.FT_PROJECTUP
                Dim Points(4) As Point
                Points(0).X = v_X
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 2
                Points(1).Y = v_Y + v_dy / 2
                Points(2).X = v_X + v_dx / 2
                Points(2).Y = v_Y + v_dy
                Points(3).X = v_X - v_dx / 2
                Points(3).Y = v_Y + v_dy
                Points(4).X = v_X - v_dx / 2
                Points(4).Y = v_Y + v_dy / 2
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_PROJECTDOWN
                Dim Points(4) As Point
                Points(0).X = v_X + v_dx / 2
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 2
                Points(1).Y = v_Y + v_dy / 2
                Points(2).X = v_X
                Points(2).Y = v_Y + v_dy
                Points(3).X = v_X - v_dx / 2
                Points(3).Y = v_Y + v_dy / 2
                Points(4).X = v_X - v_dx / 2
                Points(4).Y = v_Y
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_DIAMOND
                Dim Points(3) As Point
                Points(0).X = v_X
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 2
                Points(1).Y = v_Y + v_dy / 2
                Points(2).X = v_X
                Points(2).Y = v_Y + v_dy
                Points(3).X = v_X - v_dx / 2
                Points(3).Y = v_Y + v_dy / 2
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_CIRCLEDIAMOND
                Dim Points(3) As Point
                Points(0).X = v_X
                Points(0).Y = v_Y + v_dy / 4
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y + v_dy / 2
                Points(2).X = v_X
                Points(2).Y = v_Y + (3 * v_dy) / 4
                Points(3).X = v_X - v_dx / 4
                Points(3).Y = v_Y + v_dy / 2
                mp_DrawEllipse(v_lBorderColor, mp_oControl.MathLib.RoundDouble(v_X - v_dx / 2), v_Y, v_dx, v_dy)
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_TRIANGLEUP
                Dim Points(2) As Point
                Points(0).X = v_X
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 2
                Points(1).Y = v_Y + v_dy
                Points(2).X = v_X - v_dx / 2
                Points(2).Y = v_Y + v_dy
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_TRIANGLEDOWN
                Dim Points(2) As Point
                Points(0).X = v_X + v_dx / 2
                Points(0).Y = v_Y
                Points(1).X = v_X - v_dx / 2
                Points(1).Y = v_Y
                Points(2).X = v_X
                Points(2).Y = v_Y + v_dy
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_TRIANGLERIGHT
                Dim Points(2) As Point
                Points(0).X = v_X
                Points(0).Y = v_Y
                Points(1).X = v_X
                Points(1).Y = v_Y + v_dy
                Points(2).X = v_X + v_dx
                Points(2).Y = v_Y + v_dy / 2
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_TRIANGLELEFT
                Dim Points(2) As Point
                Points(0).X = v_X
                Points(0).Y = v_Y
                Points(1).X = v_X
                Points(1).Y = v_Y + v_dy
                Points(2).X = v_X - v_dx
                Points(2).Y = v_Y + v_dy / 2
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_CIRCLETRIANGLEUP
                Dim Points(2) As Point
                Points(0).X = v_X
                Points(0).Y = v_Y + v_dy / 4
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y + (3 * v_dy) / 4
                Points(2).X = v_X - v_dx / 4
                Points(2).Y = v_Y + (3 * v_dy) / 4
                mp_DrawEllipse(v_lBorderColor, mp_oControl.MathLib.RoundDouble(v_X - v_dx / 2), v_Y, v_dx, v_dy)
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_CIRCLETRIANGLEDOWN
                Dim Points(2) As Point
                Points(0).X = v_X - v_dx / 4
                Points(0).Y = v_Y + v_dy / 4
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y + v_dy / 4
                Points(2).X = v_X
                Points(2).Y = v_Y + (3 * v_dy) / 4
                mp_DrawEllipse(v_lBorderColor, mp_oControl.MathLib.RoundDouble(v_X - v_dx / 2), v_Y, v_dx, v_dy)
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_ARROWUP
                Dim Points(6) As Point
                Points(0).X = v_X
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 2
                Points(1).Y = v_Y + v_dy / 2
                Points(2).X = v_X + v_dx / 4
                Points(2).Y = v_Y + v_dy / 2
                Points(3).X = v_X + v_dx / 4
                Points(3).Y = v_Y + v_dy
                Points(4).X = v_X - v_dx / 4
                Points(4).Y = v_Y + v_dy
                Points(5).X = v_X - v_dx / 4
                Points(5).Y = v_Y + v_dy / 2
                Points(6).X = v_X - v_dx / 2
                Points(6).Y = v_Y + v_dy / 2
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_ARROWDOWN
                Dim Points(6) As Point
                Points(0).X = v_X - v_dx / 4
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y
                Points(2).X = v_X + v_dx / 4
                Points(2).Y = v_Y + v_dy / 2
                Points(3).X = v_X + v_dx / 2
                Points(3).Y = v_Y + v_dy / 2
                Points(4).X = v_X
                Points(4).Y = v_Y + v_dy
                Points(5).X = v_X - v_dx / 2
                Points(5).Y = v_Y + v_dy / 2
                Points(6).X = v_X - v_dx / 4
                Points(6).Y = v_Y + v_dy / 2
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_CIRCLEARROWUP
                Dim Points(6) As Point
                Points(0).X = v_X
                Points(0).Y = v_Y + v_dy / 4
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y + v_dy / 2
                Points(2).X = v_X + v_dx / 8
                Points(2).Y = v_Y + v_dy / 2
                Points(3).X = v_X + v_dx / 8
                Points(3).Y = v_Y + (3 * v_dy) / 4
                Points(4).X = v_X - v_dx / 8
                Points(4).Y = v_Y + (3 * v_dy) / 4
                Points(5).X = v_X - v_dx / 8
                Points(5).Y = v_Y + v_dy / 2
                Points(6).X = v_X - v_dx / 4
                Points(6).Y = v_Y + v_dy / 2
                mp_DrawEllipse(v_lBorderColor, mp_oControl.MathLib.RoundDouble(v_X - v_dx / 2), v_Y, v_dx, v_dy)
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_CIRCLEARROWDOWN
                Dim Points(6) As Point
                Points(0).X = v_X - v_dx / 8
                Points(0).Y = v_Y + v_dy / 4
                Points(1).X = v_X + v_dx / 8
                Points(1).Y = v_Y + v_dy / 4
                Points(2).X = v_X + v_dx / 8
                Points(2).Y = v_Y + v_dy / 2
                Points(3).X = v_X + v_dx / 4
                Points(3).Y = v_Y + v_dy / 2
                Points(4).X = v_X
                Points(4).Y = v_Y + (3 * v_dy) / 4
                Points(5).X = v_X - v_dx / 4
                Points(5).Y = v_Y + v_dy / 2
                Points(6).X = v_X - v_dx / 8
                Points(6).Y = v_Y + v_dy / 2
                mp_DrawEllipse(v_lBorderColor, mp_oControl.MathLib.RoundDouble(v_X - v_dx / 2), v_Y, v_dx, v_dy)
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_SMALLPROJECTUP
                Dim Points(4) As Point
                Points(0).X = v_X
                Points(0).Y = v_Y + v_dy / 2
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y + (3 * v_dy) / 4
                Points(2).X = v_X + v_dx / 4
                Points(2).Y = v_Y + v_dy
                Points(3).X = v_X - v_dx / 4
                Points(3).Y = v_Y + v_dy
                Points(4).X = v_X - v_dx / 4
                Points(4).Y = v_Y + (3 * v_dy) / 4
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_SMALLPROJECTDOWN
                Dim Points(4) As Point
                Points(0).X = v_X + v_dx / 4
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y + v_dy / 4
                Points(2).X = v_X
                Points(2).Y = v_Y + v_dy / 2
                Points(3).X = v_X - v_dx / 4
                Points(3).Y = v_Y + v_dy / 4
                Points(4).X = v_X - v_dx / 4
                Points(4).Y = v_Y
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_RECTANGLE
                Dim Points(3) As Point
                Points(0).X = v_X - v_dx / 8
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 8
                Points(1).Y = v_Y
                Points(2).X = v_X + v_dx / 8
                Points(2).Y = v_Y + v_dy
                Points(3).X = v_X - v_dx / 8
                Points(3).Y = v_Y + v_dy
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_SQUARE
                Dim Points(3) As Point
                Points(0).X = v_X - v_dx / 4
                Points(0).Y = v_Y + v_dx / 4
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y + v_dx / 4
                Points(2).X = v_X + v_dx / 4
                Points(2).Y = v_Y + (3 * v_dy) / 4
                Points(3).X = v_X - v_dx / 4
                Points(3).Y = v_Y + (3 * v_dy) / 4
                mp_DrawFigureAux(v_lFillColor, v_lBorderColor, Points)
            Case GRE_FIGURETYPE.FT_CIRCLE
                mp_FillEllipse(v_lFillColor, CSng(v_X - v_dx / 2), CSng(v_Y), CSng(v_dx), CSng(v_dy))
            Case Else
                Return
        End Select

    End Sub

    Private Sub mp_DrawFigureAux(ByVal BrushColor As Color, ByVal PenColor As Color, ByRef oPoints() As Point)
        DrawPolygon(BrushColor, oPoints, True)
        DrawPolygon(PenColor, oPoints, False)
    End Sub

    Private Sub mp_DrawEllipse(ByRef PenColor As Color, ByVal left As Single, ByVal Top As Single, ByVal width As Single, ByVal height As Single)
        oGraphics.DrawEllipse(Nothing, GetPen(PenColor), New Point(left + (width / 2), Top + (height / 2)), width / 2, height / 2)
    End Sub

    Private Sub mp_FillEllipse(ByRef BrushColor As Color, ByVal left As Single, ByVal Top As Single, ByVal width As Single, ByVal height As Single)
        oGraphics.DrawEllipse(GetBrush(BrushColor), Nothing, New Point(left + (width / 2), Top + (height / 2)), width / 2, height / 2)
    End Sub

    Public Sub DrawPattern(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_lColor As Color, ByVal v_lDrawStyle As GRE_PATTERN, ByVal v_iPatternFactor As Integer)
        Dim tmp As Integer
        Dim c As Integer
        Dim c1 As Integer
        Dim c2 As Integer
        Dim i1 As Integer
        Dim j1 As Integer
        Dim i2 As Integer
        Dim j2 As Integer
        If v_X1 > v_X2 Then
            tmp = v_X1
            v_X1 = v_X2
            v_X2 = tmp
        End If
        If v_Y1 > v_Y2 Then
            tmp = v_Y1
            v_Y1 = v_Y2
            v_Y2 = tmp
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_HORIZONTALLINE Or v_lDrawStyle = GRE_PATTERN.FP_CROSS Then
            For j1 = (v_Y1 + v_iPatternFactor) To v_Y2 Step v_iPatternFactor
                DrawLine(v_X1, j1, v_X2, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Next j1
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_VERTICALLINE Or v_lDrawStyle = GRE_PATTERN.FP_CROSS Then
            For j1 = (v_X1 + v_iPatternFactor) To v_X2 Step v_iPatternFactor
                DrawLine(j1, v_Y1, j1, v_Y2, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Next j1
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_UPWARDDIAGONAL Or v_lDrawStyle = GRE_PATTERN.FP_DIAGONALCROSS Then
            c1 = Int((v_Y1 + v_X1) / v_iPatternFactor + 1)
            c2 = Int((v_Y2 + v_X2) / v_iPatternFactor)
            For c = c1 To c2
                i1 = v_X1
                i2 = v_X2
                j1 = c * v_iPatternFactor - i1
                j2 = c * v_iPatternFactor - i2
                If j2 < v_Y1 Then
                    i2 = c * v_iPatternFactor - v_Y1
                    j2 = c * v_iPatternFactor - i2
                End If
                If j1 > v_Y2 Then
                    i1 = c * v_iPatternFactor - v_Y2
                    j1 = c * v_iPatternFactor - i1
                End If
                DrawLine(i1, j1, i2, j2, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False)
            Next c
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_DOWNWARDDIAGONAL Or v_lDrawStyle = GRE_PATTERN.FP_DIAGONALCROSS Then
            c1 = Int((v_Y1 - v_X2) / v_iPatternFactor + 1)
            c2 = Int((v_Y2 - v_X1) / v_iPatternFactor)
            For c = c1 To c2
                i1 = v_X1
                i2 = v_X2
                j1 = i1 + c * v_iPatternFactor
                j2 = i2 + c * v_iPatternFactor
                If j1 < v_Y1 Then
                    i1 = v_Y1 - c * v_iPatternFactor
                    j1 = i1 + c * v_iPatternFactor
                End If
                If j2 > v_Y2 Then
                    i2 = v_Y2 - c * v_iPatternFactor
                    j2 = i2 + c * v_iPatternFactor
                End If
                DrawLine(i1, j1, i2, j2, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False)
            Next c
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_LIGHT Then
            For j1 = (v_Y1 + 1) To (v_Y2 - 1)
                If j1 Mod 2 = 0 Then
                    For j2 = (v_X1 + 1) To (v_X2 - 1) Step 4
                        DrawLine(j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    Next j2
                Else
                    For j2 = (v_X1 + 3) To (v_X2 - 1) Step 4
                        DrawLine(j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    Next j2
                End If
            Next j1
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_MEDIUM Then
            For j1 = (v_Y1 + 1) To (v_Y2 - 1)
                If j1 Mod 2 = 0 Then
                    For j2 = (v_X1 + 1) To (v_X2 - 1) Step 2
                        DrawLine(j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    Next j2
                Else
                    For j2 = (v_X1 + 2) To (v_X2 - 1) Step 2
                        DrawLine(j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    Next j2
                End If
            Next j1
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_DARK Then
            For j1 = (v_Y1 + 1) To (v_Y2 - 1)
                If j1 Mod 2 = 0 Then
                    For j2 = (v_X1 + 1) To (v_X2 - 1) Step 4
                        If j2 + 3 < v_X2 Then
                            DrawLine(j2, j1, j2 + 3, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                        Else
                            DrawLine(j2, j1, v_X2, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                        End If
                    Next j2
                Else
                    DrawLine(v_X1, j1, v_X1 + 2, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    For j2 = (v_X1 + 3) To (v_X2 - 1) Step 4
                        If j2 + 3 < v_X2 Then
                            DrawLine(j2, j1, j2 + 3, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                        Else
                            DrawLine(j2, j1, v_X2, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                        End If
                    Next j2
                End If
            Next j1
        End If
    End Sub

    Public Sub DrawTextEx(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_sParam As String, ByVal v_lFlags As clsTextFlags, ByVal v_lColor As Color, ByVal v_oFont As Font, Optional ByVal v_bClip As Boolean = True)
        Dim oTypeFace As New Typeface(v_oFont.FamilyName)
        Dim oFormattedText As New FormattedText(v_sParam, mp_oControl.Culture, FlowDirection.LeftToRight, oTypeFace, v_oFont.WPFFontSize, New SolidColorBrush(v_lColor))
        Dim X As Integer = 0
        Dim Y As Integer = 0
        Select Case v_lFlags.HorizontalAlignment
            Case GRE_HORIZONTALALIGNMENT.HAL_LEFT
                X = System.Convert.ToDouble(v_X1)
            Case GRE_HORIZONTALALIGNMENT.HAL_CENTER
                X = System.Convert.ToDouble(((v_X2 - v_X1) - oFormattedText.Width) / 2) + v_X1
            Case GRE_HORIZONTALALIGNMENT.HAL_RIGHT
                X = System.Convert.ToDouble(v_X2 - oFormattedText.Width)
        End Select
        Select Case v_lFlags.VerticalAlignment
            Case GRE_VERTICALALIGNMENT.VAL_TOP
                Y = System.Convert.ToDouble(v_Y1)
            Case GRE_VERTICALALIGNMENT.VAL_CENTER
                Y = System.Convert.ToDouble(((v_Y2 - v_Y1) - oFormattedText.Height) / 2) + v_Y1
            Case GRE_VERTICALALIGNMENT.VAL_BOTTOM
                Y = System.Convert.ToDouble(v_Y2 - oFormattedText.Height)
        End Select
        oFormattedText.SetFontWeight(v_oFont.FontWeight)
        oGraphics.DrawText(oFormattedText, New Point(X, Y))
        If v_sParam.Length > 0 Then
            mp_oTextFinalLayout.X = X
            mp_oTextFinalLayout.Y = Y
            mp_oTextFinalLayout.Width = oFormattedText.Width + mp_oControl.mp_lStrWidth("W", v_oFont)
            mp_oTextFinalLayout.Height = oFormattedText.Height
        End If
    End Sub

    Public Sub DrawAlignedText(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal v_sParam As String, ByVal v_yHPos As GRE_HORIZONTALALIGNMENT, ByVal v_yVPos As GRE_VERTICALALIGNMENT, ByVal v_lColor As Color, ByVal v_oFont As Font)
        DrawAlignedText(v_lLeft, v_lTop, v_lRight, v_lBottom, v_sParam, v_yHPos, v_yVPos, v_lColor, v_oFont, True)
    End Sub

    Public Sub DrawAlignedText(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal v_sParam As String, ByVal v_yHPos As GRE_HORIZONTALALIGNMENT, ByVal v_yVPos As GRE_VERTICALALIGNMENT, ByVal v_lColor As Color, ByVal v_oFont As Font, ByVal v_bClip As Boolean)
        Dim oTypeFace As New Typeface(v_oFont.FamilyName)
        Dim oFormattedText As New FormattedText(v_sParam, mp_oControl.Culture, FlowDirection.LeftToRight, oTypeFace, v_oFont.WPFFontSize, New SolidColorBrush(v_lColor))
        Dim X As Integer = 0
        Dim Y As Integer = 0
        Select Case v_yHPos
            Case GRE_HORIZONTALALIGNMENT.HAL_LEFT
                X = System.Convert.ToDouble(v_lLeft)
            Case GRE_HORIZONTALALIGNMENT.HAL_CENTER
                X = System.Convert.ToDouble(((v_lRight - v_lLeft) - oFormattedText.Width) / 2) + v_lLeft
            Case GRE_HORIZONTALALIGNMENT.HAL_RIGHT
                X = System.Convert.ToDouble(v_lRight - oFormattedText.Width)
        End Select
        Select Case v_yVPos
            Case GRE_VERTICALALIGNMENT.VAL_TOP
                Y = System.Convert.ToDouble(v_lTop)
            Case GRE_VERTICALALIGNMENT.VAL_CENTER
                Y = System.Convert.ToDouble(((v_lBottom - v_lTop) - oFormattedText.Height) / 2) + v_lTop
            Case GRE_VERTICALALIGNMENT.VAL_BOTTOM
                Y = System.Convert.ToDouble(v_lBottom - oFormattedText.Height)
        End Select
        oFormattedText.SetFontWeight(v_oFont.FontWeight)
        oGraphics.DrawText(oFormattedText, New Point(X, Y))
        If v_sParam.Length > 0 Then
            mp_oTextFinalLayout.X = X
            mp_oTextFinalLayout.Y = Y
            mp_oTextFinalLayout.Width = oFormattedText.Width + mp_oControl.mp_lStrWidth("W", v_oFont)
            mp_oTextFinalLayout.Height = oFormattedText.Height
        End If
    End Sub

    Public Sub ClipRegion(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_bStore As Boolean)
        If (mp_bEnableClipRegions = False) Then
            Return
        End If
        If mp_bRequiresPop = True Then
            oGraphics.Pop()
            mp_bRequiresPop = False
        End If
        CorrectRectCoords(v_X1, v_Y1, v_X2, v_Y2)
        Dim oRectangle As New RectangleGeometry(New Rect(v_X1, v_Y1, v_X2 - v_X1 + 1, v_Y2 - v_Y1 + 1))
        oRectangle.Freeze()
        If v_bStore = True Then
            mp_udtPreviousClipRegion.lLeft = v_X1
            mp_udtPreviousClipRegion.lRight = v_X2
            mp_udtPreviousClipRegion.lTop = v_Y1
            mp_udtPreviousClipRegion.lBottom = v_Y2
        End If
        oGraphics.PushClip(oRectangle)
        mp_bRequiresPop = True

    End Sub

    Public Sub RestorePreviousClipRegion()
        If (mp_bEnableClipRegions = False) Then
            Return
        End If
        ClipRegion(mp_udtPreviousClipRegion.lLeft, mp_udtPreviousClipRegion.lTop, mp_udtPreviousClipRegion.lRight, mp_udtPreviousClipRegion.lBottom, False)
    End Sub

    Public Sub ClearClipRegion()
        If mp_bRequiresPop = True Then
            oGraphics.Pop()
            mp_bRequiresPop = False
        End If
    End Sub

    Public Sub TileImageHorizontal(ByVal ImageHandle As Image, ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_bTransparent As Boolean)
        Dim X As Integer
        Dim lImageWidth As Integer
        Dim lImageHeight As Integer
        lImageHeight = ImageHandle.Source.Height
        lImageWidth = ImageHandle.Source.Width
        Do While X < (v_X2 - v_X1)
            If (X + lImageWidth) > (v_X2 - v_X1) Then
                PaintImage(ImageHandle, v_X2 - lImageWidth, v_Y1, v_X2, v_Y1 + lImageHeight, 0, 0, v_bTransparent)
            Else
                PaintImage(ImageHandle, v_X1 + X, v_Y1, v_X1 + X + lImageWidth, v_Y1 + lImageHeight, 0, 0, v_bTransparent)
            End If
            X = X + lImageWidth
        Loop
    End Sub

    Public Sub PaintImage(ByVal Image As Image, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal xOrigin As Integer, ByVal yOrigin As Integer, ByVal bUseMask As Boolean)
        'Dim oImage As New Image
        If xOrigin <> 0 Or yOrigin <> 0 Then
            'Dim oCroppedImage As New CroppedBitmap(Image, New Int32Rect(xOrigin, yOrigin, X2 - xOrigin, Y2 - yOrigin))
            'oCroppedImage.Source = Image
            'oImage.Source = oCroppedImage
        Else
            'oImage.Source = Image
        End If
        oGraphics.DrawImage(Image.Source, New Rect(X1, Y1, X2 - X1, Y2 - Y1))
    End Sub

    Public Sub DrawImage(ByRef v_oImage As Image, ByRef v_yHorizontalAlignment As GRE_HORIZONTALALIGNMENT, ByRef v_yVerticalAlignment As GRE_VERTICALALIGNMENT, ByVal v_lImageXMargin As Integer, ByVal v_lImageYMargin As Integer, ByRef v_lLeft As Integer, ByRef v_lRight As Integer, ByRef v_lTop As Integer, ByRef v_lBottom As Integer, ByVal v_bTransparent As Boolean)
        Dim bDrawImage As Boolean
        Dim bHorizontalSmall As Boolean
        Dim bVerticalSmall As Boolean
        Dim XOrigin As Integer
        Dim YOrigin As Integer
        Dim xDest As Integer
        Dim yDest As Integer
        Dim lxWidth As Integer
        Dim lyHeight As Integer
        Dim lImageHeight As Integer
        Dim lImageWidth As Integer
        If (v_oImage Is Nothing) Then
            Return
        End If
        lImageHeight = v_oImage.Source.Height
        lImageWidth = v_oImage.Source.Width
        If v_yHorizontalAlignment = GRE_HORIZONTALALIGNMENT.HAL_CENTER Then
            v_lImageXMargin = 0
        End If
        If v_yVerticalAlignment = GRE_VERTICALALIGNMENT.VAL_CENTER Then
            v_lImageYMargin = 0
        End If
        bDrawImage = True
        If (v_lRight - v_lLeft) < (lImageWidth + v_lImageXMargin) Then
            lxWidth = v_lRight - v_lLeft - v_lImageXMargin
            If lxWidth <= 0 Then bDrawImage = False
            bHorizontalSmall = True
        Else
            lxWidth = lImageWidth
            bHorizontalSmall = False
        End If
        If (v_lBottom - v_lTop) < (lImageHeight + v_lImageYMargin) Then
            lyHeight = v_lBottom - v_lTop - v_lImageYMargin
            If lyHeight <= 0 Then bDrawImage = False
            bVerticalSmall = True
        Else
            lyHeight = lImageHeight
            bVerticalSmall = False
        End If
        If bHorizontalSmall = False Then
            Select Case v_yHorizontalAlignment
                Case GRE_HORIZONTALALIGNMENT.HAL_LEFT
                    xDest = v_lLeft + v_lImageXMargin
                Case GRE_HORIZONTALALIGNMENT.HAL_CENTER
                    xDest = ((v_lRight - v_lLeft) - lImageWidth) / 2 + v_lLeft
                Case GRE_HORIZONTALALIGNMENT.HAL_RIGHT
                    xDest = v_lRight - lImageWidth - v_lImageXMargin
            End Select
            XOrigin = 0
        Else
            Select Case v_yHorizontalAlignment
                Case GRE_HORIZONTALALIGNMENT.HAL_LEFT
                    XOrigin = 0
                    xDest = v_lLeft + v_lImageXMargin
                Case GRE_HORIZONTALALIGNMENT.HAL_CENTER
                    XOrigin = (lImageWidth - lxWidth) / 2
                    xDest = v_lLeft
                Case GRE_HORIZONTALALIGNMENT.HAL_RIGHT
                    XOrigin = lImageWidth - lxWidth
                    xDest = v_lRight - lxWidth - v_lImageXMargin
            End Select
        End If
        If bVerticalSmall = False Then
            Select Case v_yVerticalAlignment
                Case GRE_VERTICALALIGNMENT.VAL_TOP
                    yDest = v_lTop + v_lImageYMargin
                Case GRE_VERTICALALIGNMENT.VAL_CENTER
                    yDest = ((v_lBottom - v_lTop) - lImageHeight) / 2 + v_lTop
                Case GRE_VERTICALALIGNMENT.VAL_BOTTOM
                    yDest = v_lBottom - lImageHeight - v_lImageYMargin
            End Select
            YOrigin = 0
        Else
            Select Case v_yVerticalAlignment
                Case GRE_VERTICALALIGNMENT.VAL_TOP
                    YOrigin = 0
                    yDest = v_lTop + v_lImageYMargin
                Case GRE_VERTICALALIGNMENT.VAL_CENTER
                    YOrigin = (lImageHeight - lyHeight) / 2
                    yDest = v_lTop
                Case GRE_VERTICALALIGNMENT.VAL_BOTTOM
                    YOrigin = lImageHeight - lyHeight
                    yDest = v_lBottom - lyHeight - v_lImageYMargin
            End Select
        End If
        If bDrawImage = True Then
            PaintImage(v_oImage, xDest, yDest, xDest + lxWidth, yDest + lyHeight, XOrigin, YOrigin, v_bTransparent)
        End If
    End Sub

    Public Sub DrawFocusRectangle(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer)
        DrawLine(v_X1, v_Y1, v_X2, v_Y2, GRE_LINETYPE.LT_BORDER, Colors.Black, GRE_LINEDRAWSTYLE.LDS_DOT)
    End Sub

    Public Sub GradientFill(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal clrStartColor As Color, ByVal clrEndColor As Color, ByVal iGradientType As GRE_GRADIENTFILLMODE)
        If (v_X2 - v_X1) <= 0 Then
            Return
        End If
        If (v_Y2 - v_Y1) <= 0 Then
            Return
        End If
        Dim mp_ucBrush As LinearGradientBrush = Nothing
        If (iGradientType = GRE_GRADIENTFILLMODE.GDT_VERTICAL) Then
            mp_ucBrush = New LinearGradientBrush(clrStartColor, clrEndColor, 90.0)
        ElseIf (iGradientType = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL) Then
            mp_ucBrush = New LinearGradientBrush(clrStartColor, clrEndColor, 0.0)
        End If
        mp_ucBrush.Freeze()
        oGraphics.DrawRectangle(mp_ucBrush, Nothing, New Rect(v_X1, v_Y1, v_X2 - v_X1 + 1, v_Y2 - v_Y1 + 1))
    End Sub

    Public Sub HatchFill(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal clrForeColor As Color, ByVal clrBackColor As Color, ByVal yHatchStyle As GRE_HATCHSTYLE)
        Dim oBrush As New DrawingBrush
        Dim oHatchGroup As New GeometryGroup()
        Dim oHatchCtrlGroup As New GeometryGroup()
        Dim lWidth As Integer = 0
        Dim lHeight As Integer = 0
        Dim yType As T_HATCHTYPE = T_HATCHTYPE.HT_LINE
        Dim bAliased As Boolean = True
        Dim iBuff As Integer = 0
        If (v_X2 - v_X1) <= 0 Then
            iBuff = v_X1
            v_X1 = v_X2
            v_X2 = iBuff
        End If
        If (v_Y2 - v_Y1) <= 0 Then
            iBuff = v_Y1
            v_Y1 = v_Y2
            v_Y2 = iBuff
        End If
        Select Case yHatchStyle
            Case GRE_HATCHSTYLE.HS_HORIZONTAL
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(0, 0, 7, 0))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_VERTICAL
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(0, 0, 0, 7))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_FORWARDDIAGONAL
                lWidth = 16
                lHeight = 16
                oHatchGroup.Children.Add(mp_GLine(0, 12, 3, 15))
                oHatchGroup.Children.Add(mp_GLine(0, 4, 11, 15))
                oHatchGroup.Children.Add(mp_GLine(4, 0, 15, 11))
                oHatchGroup.Children.Add(mp_GLine(12, 0, 15, 3))
                System.Diagnostics.Debug.Write(oHatchGroup)
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_BACKWARDDIAGONAL
                lWidth = 16
                lHeight = 16
                oHatchGroup.Children.Add(mp_GLine(0, 12, 3, 15))
                oHatchGroup.Children.Add(mp_GLine(0, 4, 11, 15))
                oHatchGroup.Children.Add(mp_GLine(4, 0, 15, 11))
                oHatchGroup.Children.Add(mp_GLine(12, 0, 15, 3))
                oHatchGroup.Transform = New RotateTransform(90, 8, 8)
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_LARGEGRID
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(0, 0, 7, 0))
                oHatchGroup.Children.Add(mp_GLine(0, 0, 0, 7))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_DIAGONALCROSS
                lWidth = 7
                lHeight = 7
                oHatchGroup.Children.Add(mp_GRect(1, 1, 5, 5))
                oHatchGroup.Transform = New RotateTransform(45, 3, 3)
                yType = T_HATCHTYPE.HT_LINE
                bAliased = False
            Case GRE_HATCHSTYLE.HS_PERCENT05
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 4))
                oHatchGroup.Children.Add(mp_GPoint(4, 0))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_PERCENT10
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                oHatchGroup.Children.Add(mp_GPoint(4, 0))
                oHatchGroup.Children.Add(mp_GPoint(0, 6))
                oHatchGroup.Children.Add(mp_GPoint(4, 4))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_PERCENT20
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_PERCENT25
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                oHatchGroup.Children.Add(mp_GPoint(2, 1))
                oHatchGroup.Children.Add(mp_GPoint(2, 3))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_PERCENT30
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                oHatchGroup.Children.Add(mp_GPoint(2, 0))
                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_PERCENT40
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(2, 0))
                oHatchGroup.Children.Add(mp_GPoint(4, 0))
                oHatchGroup.Children.Add(mp_GPoint(6, 0))

                oHatchGroup.Children.Add(mp_GPoint(3, 1))
                oHatchGroup.Children.Add(mp_GPoint(5, 1))
                oHatchGroup.Children.Add(mp_GPoint(7, 1))

                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(4, 2))
                oHatchGroup.Children.Add(mp_GPoint(6, 2))

                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                oHatchGroup.Children.Add(mp_GPoint(5, 3))
                oHatchGroup.Children.Add(mp_GPoint(7, 3))

                oHatchGroup.Children.Add(mp_GPoint(0, 4))
                oHatchGroup.Children.Add(mp_GPoint(2, 4))
                oHatchGroup.Children.Add(mp_GPoint(4, 4))
                oHatchGroup.Children.Add(mp_GPoint(6, 4))
                oHatchGroup.Children.Add(mp_GPoint(1, 5))
                oHatchGroup.Children.Add(mp_GPoint(3, 5))
                oHatchGroup.Children.Add(mp_GPoint(7, 5))
                oHatchGroup.Children.Add(mp_GPoint(0, 6))
                oHatchGroup.Children.Add(mp_GPoint(2, 6))
                oHatchGroup.Children.Add(mp_GPoint(4, 6))
                oHatchGroup.Children.Add(mp_GPoint(6, 6))
                oHatchGroup.Children.Add(mp_GPoint(1, 7))
                oHatchGroup.Children.Add(mp_GPoint(3, 7))
                oHatchGroup.Children.Add(mp_GPoint(5, 7))
                oHatchGroup.Children.Add(mp_GPoint(7, 7))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_PERCENT50
                lWidth = 2
                lHeight = 2
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_PERCENT60
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(2, 0))
                oHatchGroup.Children.Add(mp_GPoint(3, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                oHatchGroup.Children.Add(mp_GPoint(3, 1))
                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                oHatchGroup.Children.Add(mp_GPoint(1, 2))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_PERCENT70
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(1, 0))
                oHatchGroup.Children.Add(mp_GPoint(2, 0))
                oHatchGroup.Children.Add(mp_GPoint(3, 0))
                oHatchGroup.Children.Add(mp_GPoint(0, 1))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                oHatchGroup.Children.Add(mp_GPoint(3, 1))
                oHatchGroup.Children.Add(mp_GPoint(1, 2))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(3, 2))
                oHatchGroup.Children.Add(mp_GPoint(0, 3))
                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_PERCENT75
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(2, 0))
                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                InvertColors(clrForeColor, clrBackColor)
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_PERCENT80
                lWidth = 8
                lHeight = 7
                oHatchGroup.Children.Add(mp_GPoint(3, 0))
                oHatchGroup.Children.Add(mp_GPoint(3, 4))
                oHatchGroup.Children.Add(mp_GPoint(7, 2))
                oHatchGroup.Children.Add(mp_GPoint(7, 6))
                InvertColors(clrForeColor, clrBackColor)
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_PERCENT90
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 5))
                oHatchGroup.Children.Add(mp_GPoint(4, 1))
                InvertColors(clrForeColor, clrBackColor)
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_LIGHTDOWNWARDDIAGONAL
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_LIGHTUPWARDDIAGONAL
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 3))
                oHatchGroup.Children.Add(mp_GPoint(1, 2))
                oHatchGroup.Children.Add(mp_GPoint(2, 1))
                oHatchGroup.Children.Add(mp_GPoint(3, 0))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_DARKDOWNWARDDIAGONAL
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 3))
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                oHatchGroup.Children.Add(mp_GPoint(1, 0))
                oHatchGroup.Children.Add(mp_GPoint(2, 1))
                oHatchGroup.Children.Add(mp_GPoint(3, 2))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_DARKUPWARDDIAGONAL
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(0, 1))
                oHatchGroup.Children.Add(mp_GPoint(1, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(3, 1))
                oHatchGroup.Children.Add(mp_GPoint(2, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 2))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_WIDEDOWNWARDDIAGONAL
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(3, 0, 5, 0))
                oHatchGroup.Children.Add(mp_GLine(4, 1, 6, 1))
                oHatchGroup.Children.Add(mp_GLine(5, 2, 7, 2))
                oHatchGroup.Children.Add(mp_GPoint(0, 3))
                oHatchGroup.Children.Add(mp_GLine(6, 3, 7, 3))
                oHatchGroup.Children.Add(mp_GLine(0, 4, 1, 4))
                oHatchGroup.Children.Add(mp_GPoint(7, 4))
                oHatchGroup.Children.Add(mp_GLine(0, 5, 2, 5))
                oHatchGroup.Children.Add(mp_GLine(1, 6, 3, 6))
                oHatchGroup.Children.Add(mp_GLine(2, 7, 4, 7))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_WIDEUPWARDDIAGONAL
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(4, 0, 6, 0))
                oHatchGroup.Children.Add(mp_GLine(3, 1, 5, 1))
                oHatchGroup.Children.Add(mp_GLine(2, 2, 4, 2))
                oHatchGroup.Children.Add(mp_GLine(1, 3, 3, 3))
                oHatchGroup.Children.Add(mp_GLine(0, 4, 2, 4))
                oHatchGroup.Children.Add(mp_GLine(0, 5, 1, 5))
                oHatchGroup.Children.Add(mp_GPoint(7, 5))
                oHatchGroup.Children.Add(mp_GPoint(0, 6))
                oHatchGroup.Children.Add(mp_GLine(6, 6, 7, 6))
                oHatchGroup.Children.Add(mp_GLine(5, 7, 7, 7))
                oHatchCtrlGroup.Children.Add(mp_GLine(0, 0, 7, 7))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_LIGHTVERTICAL
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GLine(0, 0, 0, 3))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_LIGHTHORIZONTAL
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GLine(0, 0, 3, 0))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_NARROWVERTICAL
                lWidth = 2
                lHeight = 2
                oHatchGroup.Children.Add(mp_GLine(1, 0, 1, 1))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_NARROWHORIZONTAL
                lWidth = 2
                lHeight = 2
                oHatchGroup.Children.Add(mp_GLine(0, 1, 1, 1))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_DARKVERTICAL
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GLine(0, 0, 0, 3))
                oHatchGroup.Children.Add(mp_GLine(1, 0, 1, 3))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_DARKHORIZONTAL
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GLine(0, 0, 3, 0))
                oHatchGroup.Children.Add(mp_GLine(0, 1, 3, 1))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_DASHEDDOWNWARDDIAGONAL
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(2, 4))
                oHatchGroup.Children.Add(mp_GPoint(3, 5))
                oHatchGroup.Children.Add(mp_GPoint(4, 2))
                oHatchGroup.Children.Add(mp_GPoint(5, 3))
                oHatchGroup.Children.Add(mp_GPoint(6, 4))
                oHatchGroup.Children.Add(mp_GPoint(7, 5))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_DASHEDUPWARDDIAGONAL
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 7))
                oHatchGroup.Children.Add(mp_GPoint(1, 6))
                oHatchGroup.Children.Add(mp_GPoint(2, 5))
                oHatchGroup.Children.Add(mp_GPoint(3, 4))
                oHatchGroup.Children.Add(mp_GPoint(4, 7))
                oHatchGroup.Children.Add(mp_GPoint(5, 6))
                oHatchGroup.Children.Add(mp_GPoint(6, 5))
                oHatchGroup.Children.Add(mp_GPoint(7, 4))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_DASHEDHORIZONTAL
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(4, 0, 7, 0))
                oHatchGroup.Children.Add(mp_GLine(0, 4, 3, 4))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_DASHEDVERTICAL
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(0, 0, 0, 1))
                oHatchGroup.Children.Add(mp_GLine(0, 6, 0, 7))
                oHatchGroup.Children.Add(mp_GLine(4, 2, 4, 5))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_SMALLCONFETTI
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(4, 1))
                oHatchGroup.Children.Add(mp_GPoint(1, 2))
                oHatchGroup.Children.Add(mp_GPoint(6, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 4))
                oHatchGroup.Children.Add(mp_GPoint(7, 5))
                oHatchGroup.Children.Add(mp_GPoint(2, 6))
                oHatchGroup.Children.Add(mp_GPoint(5, 7))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_LARGECONFETTI
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(0, 1, 0, 2))
                oHatchGroup.Children.Add(mp_GLine(0, 6, 0, 7))
                oHatchGroup.Children.Add(mp_GLine(1, 6, 1, 7))
                oHatchGroup.Children.Add(mp_GLine(2, 2, 2, 3))
                oHatchGroup.Children.Add(mp_GLine(3, 2, 3, 3))
                oHatchGroup.Children.Add(mp_GLine(3, 5, 3, 6))
                oHatchGroup.Children.Add(mp_GLine(4, 0, 4, 1))
                oHatchGroup.Children.Add(mp_GLine(4, 5, 4, 6))
                oHatchGroup.Children.Add(mp_GLine(5, 0, 5, 1))
                oHatchGroup.Children.Add(mp_GLine(6, 4, 6, 5))
                oHatchGroup.Children.Add(mp_GLine(7, 1, 7, 2))
                oHatchGroup.Children.Add(mp_GLine(7, 4, 7, 5))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_ZIGZAG
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GLine(3, 3, 4, 3))
                oHatchGroup.Children.Add(mp_GPoint(5, 2))
                oHatchGroup.Children.Add(mp_GPoint(6, 1))
                oHatchGroup.Children.Add(mp_GPoint(7, 0))

                oHatchGroup.Children.Add(mp_GPoint(0, 4))
                oHatchGroup.Children.Add(mp_GPoint(1, 5))
                oHatchGroup.Children.Add(mp_GPoint(2, 6))
                oHatchGroup.Children.Add(mp_GLine(3, 7, 4, 7))
                oHatchGroup.Children.Add(mp_GPoint(5, 6))
                oHatchGroup.Children.Add(mp_GPoint(6, 5))
                oHatchGroup.Children.Add(mp_GPoint(7, 4))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_WAVE
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(2, 0))
                oHatchGroup.Children.Add(mp_GPoint(5, 0))
                oHatchGroup.Children.Add(mp_GPoint(7, 0))
                oHatchGroup.Children.Add(mp_GLine(0, 2, 1, 2))

                oHatchGroup.Children.Add(mp_GLine(3, 4, 4, 4))
                oHatchGroup.Children.Add(mp_GPoint(2, 4))
                oHatchGroup.Children.Add(mp_GPoint(5, 4))
                oHatchGroup.Children.Add(mp_GPoint(7, 4))

                oHatchGroup.Children.Add(mp_GLine(0, 6, 1, 6))
                oHatchGroup.Children.Add(mp_GLine(3, 8, 4, 8))
                oHatchCtrlGroup.Children.Add(mp_GPoint(0, 0))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_DIAGONALBRICK
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 7))
                oHatchGroup.Children.Add(mp_GPoint(1, 6))
                oHatchGroup.Children.Add(mp_GPoint(2, 5))
                oHatchGroup.Children.Add(mp_GPoint(3, 4))
                oHatchGroup.Children.Add(mp_GPoint(4, 3))
                oHatchGroup.Children.Add(mp_GPoint(5, 2))
                oHatchGroup.Children.Add(mp_GPoint(6, 1))
                oHatchGroup.Children.Add(mp_GPoint(7, 0))
                oHatchGroup.Children.Add(mp_GPoint(4, 4))
                oHatchGroup.Children.Add(mp_GPoint(5, 5))
                oHatchGroup.Children.Add(mp_GPoint(6, 6))
                oHatchGroup.Children.Add(mp_GPoint(7, 7))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_HORIZONTALBRICK
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(4, 0))
                oHatchGroup.Children.Add(mp_GPoint(4, 1))
                oHatchGroup.Children.Add(mp_GLine(0, 1, 0, 5))
                oHatchGroup.Children.Add(mp_GPoint(4, 6))
                oHatchGroup.Children.Add(mp_GPoint(4, 7))

                oHatchGroup.Children.Add(mp_GLine(1, 2, 7, 2))
                oHatchGroup.Children.Add(mp_GLine(1, 6, 7, 6))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_WEAVE
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(0, 4))
                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 1))
                oHatchGroup.Children.Add(mp_GPoint(4, 0))
                oHatchGroup.Children.Add(mp_GPoint(5, 1))
                oHatchGroup.Children.Add(mp_GPoint(6, 2))
                oHatchGroup.Children.Add(mp_GPoint(7, 3))
                oHatchGroup.Children.Add(mp_GPoint(5, 3))
                oHatchGroup.Children.Add(mp_GPoint(4, 4))
                oHatchGroup.Children.Add(mp_GPoint(3, 5))
                oHatchGroup.Children.Add(mp_GPoint(2, 6))
                oHatchGroup.Children.Add(mp_GPoint(1, 7))
                oHatchGroup.Children.Add(mp_GPoint(3, 7))
                oHatchGroup.Children.Add(mp_GPoint(5, 5))
                oHatchGroup.Children.Add(mp_GPoint(6, 6))
                oHatchGroup.Children.Add(mp_GPoint(7, 7))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_PLAID
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(0, 0, 3, 0))
                oHatchGroup.Children.Add(mp_GLine(0, 1, 3, 1))

                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(4, 2))
                oHatchGroup.Children.Add(mp_GPoint(6, 2))

                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                oHatchGroup.Children.Add(mp_GPoint(5, 3))
                oHatchGroup.Children.Add(mp_GPoint(7, 3))


                oHatchGroup.Children.Add(mp_GPoint(0, 4))
                oHatchGroup.Children.Add(mp_GPoint(2, 4))
                oHatchGroup.Children.Add(mp_GPoint(4, 4))
                oHatchGroup.Children.Add(mp_GPoint(6, 4))

                oHatchGroup.Children.Add(mp_GPoint(1, 5))
                oHatchGroup.Children.Add(mp_GPoint(3, 5))
                oHatchGroup.Children.Add(mp_GPoint(5, 5))
                oHatchGroup.Children.Add(mp_GPoint(7, 5))

                oHatchGroup.Children.Add(mp_GLine(0, 6, 3, 6))
                oHatchGroup.Children.Add(mp_GLine(0, 7, 3, 7))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_DIVOT
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 1))
                oHatchGroup.Children.Add(mp_GPoint(0, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 5))
                oHatchGroup.Children.Add(mp_GPoint(3, 7))
                oHatchGroup.Children.Add(mp_GPoint(4, 6))
                oHatchGroup.Children.Add(mp_GPoint(7, 2))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_DOTTEDGRID
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(1, 6))
                oHatchGroup.Children.Add(mp_GPoint(3, 6))
                oHatchGroup.Children.Add(mp_GPoint(5, 6))
                oHatchGroup.Children.Add(mp_GPoint(7, 6))
                oHatchGroup.Children.Add(mp_GPoint(7, 4))
                oHatchGroup.Children.Add(mp_GPoint(7, 2))
                oHatchGroup.Children.Add(mp_GPoint(7, 0))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_DOTTEDDIAMOND
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(4, 4))
                oHatchGroup.Children.Add(mp_GPoint(6, 6))
                oHatchGroup.Children.Add(mp_GPoint(2, 6))
                oHatchGroup.Children.Add(mp_GPoint(6, 2))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_SHINGLE
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 3))
                oHatchGroup.Children.Add(mp_GPoint(1, 4))
                oHatchGroup.Children.Add(mp_GPoint(2, 5))
                oHatchGroup.Children.Add(mp_GPoint(3, 5))
                oHatchGroup.Children.Add(mp_GPoint(4, 6))
                oHatchGroup.Children.Add(mp_GPoint(5, 6))
                oHatchGroup.Children.Add(mp_GPoint(6, 7))
                oHatchGroup.Children.Add(mp_GPoint(4, 4))
                oHatchGroup.Children.Add(mp_GPoint(5, 3))
                oHatchGroup.Children.Add(mp_GPoint(6, 2))
                oHatchGroup.Children.Add(mp_GPoint(7, 2))
                oHatchGroup.Children.Add(mp_GPoint(7, 0))
                oHatchGroup.Children.Add(mp_GPoint(7, 1))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_TRELLIS
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GLine(0, 0, 3, 0))
                oHatchGroup.Children.Add(mp_GLine(1, 1, 2, 1))
                oHatchGroup.Children.Add(mp_GLine(0, 2, 3, 2))
                oHatchGroup.Children.Add(mp_GPoint(0, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_SPHERE
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(1, 0, 3, 0))
                oHatchGroup.Children.Add(mp_GLine(1, 1, 3, 1))
                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                oHatchGroup.Children.Add(mp_GPoint(4, 2))
                oHatchGroup.Children.Add(mp_GLine(1, 3, 2, 3))
                oHatchGroup.Children.Add(mp_GPoint(0, 6))
                oHatchGroup.Children.Add(mp_GPoint(4, 6))
                oHatchGroup.Children.Add(mp_GLine(1, 7, 3, 7))
                oHatchGroup.Children.Add(mp_GLine(5, 7, 6, 7))
                oHatchGroup.Children.Add(mp_GLine(5, 3, 7, 3))
                oHatchGroup.Children.Add(mp_GLine(5, 4, 7, 4))
                oHatchGroup.Children.Add(mp_GLine(5, 5, 7, 5))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                InvertColors(clrForeColor, clrBackColor)
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_SMALLGRID
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GLine(0, 0, 3, 0))
                oHatchGroup.Children.Add(mp_GLine(0, 0, 0, 3))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_SMALLCHECKERBOARD
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GRect(0, 0, 2, 2))
                oHatchGroup.Children.Add(mp_GRect(2, 2, 2, 2))
                yType = T_HATCHTYPE.HT_RECTANGLE
            Case GRE_HATCHSTYLE.HS_LARGECHECKERBOARD
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GRect(0, 0, 4, 4))
                oHatchGroup.Children.Add(mp_GRect(4, 4, 4, 4))
                yType = T_HATCHTYPE.HT_RECTANGLE
            Case GRE_HATCHSTYLE.HS_OUTLINEDDIAMOND
                lWidth = 8
                lHeight = 8

                oHatchGroup.Children.Add(mp_GPoint(0, 4))
                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(3, 1))
                oHatchGroup.Children.Add(mp_GPoint(4, 0))

                oHatchGroup.Children.Add(mp_GPoint(5, 1))
                oHatchGroup.Children.Add(mp_GPoint(6, 2))
                oHatchGroup.Children.Add(mp_GPoint(7, 3))

                oHatchGroup.Children.Add(mp_GPoint(7, 5))
                oHatchGroup.Children.Add(mp_GPoint(6, 6))
                oHatchGroup.Children.Add(mp_GPoint(5, 7))

                oHatchGroup.Children.Add(mp_GPoint(3, 7))
                oHatchGroup.Children.Add(mp_GPoint(2, 6))
                oHatchGroup.Children.Add(mp_GPoint(1, 5))
                yType = T_HATCHTYPE.HT_LINE
            Case GRE_HATCHSTYLE.HS_SOLIDDIAMOND
                lWidth = 7
                lHeight = 7
                oHatchGroup.Children.Add(mp_GRect(1, 1, 5, 5))
                oHatchGroup.Transform = New RotateTransform(45, 3, 3)
                yType = T_HATCHTYPE.HT_RECTANGLE
        End Select
        Dim oBackgroundSquare As New GeometryDrawing(New SolidColorBrush(clrBackColor), Nothing, New RectangleGeometry(New Rect(0, 0, lWidth, lHeight)))
        Dim oHatchBrush As New SolidColorBrush(clrForeColor)
        Dim oHatchPen As New Pen(oHatchBrush, 1)
        oHatchBrush.Freeze()
        oHatchPen.Freeze()
        Dim oHatchCtrlBrush As New SolidColorBrush(Colors.Red)
        Dim oHatchCtrlPen As New Pen(oHatchCtrlBrush, 1)
        oHatchCtrlBrush.Freeze()
        oHatchCtrlPen.Freeze()
        Dim oHatch As GeometryDrawing = Nothing
        Dim oHatchCtrl As GeometryDrawing = Nothing
        Select Case yType
            Case T_HATCHTYPE.HT_RECTANGLE
                oHatch = New GeometryDrawing(oHatchBrush, Nothing, oHatchGroup)
                If oHatchCtrlGroup.Children.Count > 0 Then
                    oHatchCtrl = New GeometryDrawing(oHatchCtrlBrush, Nothing, oHatchCtrlGroup)
                End If
            Case T_HATCHTYPE.HT_LINE
                oHatch = New GeometryDrawing(Nothing, oHatchPen, oHatchGroup)
                If oHatchCtrlGroup.Children.Count > 0 Then
                    oHatchCtrl = New GeometryDrawing(Nothing, oHatchCtrlPen, oHatchCtrlGroup)
                End If
        End Select
        Dim oDrawingGroup As New DrawingGroup
        If bAliased = True Then
            oDrawingGroup.SetValue(RenderOptions.EdgeModeProperty, EdgeMode.Aliased)
        End If
        If Not oHatchCtrl Is Nothing Then
            oDrawingGroup.Children.Add(oHatchCtrl)
        End If
        oDrawingGroup.Children.Add(oBackgroundSquare)
        oDrawingGroup.Children.Add(oHatch)
        oBrush.Drawing = oDrawingGroup
        oBrush.Stretch = Stretch.None
        oBrush.ViewportUnits = BrushMappingMode.Absolute
        oBrush.Viewport = New Rect(0, 0, lWidth, lHeight)
        oBrush.TileMode = TileMode.Tile
        oBrush.Freeze()
        oGraphics.DrawRectangle(oBrush, Nothing, New Rect(v_X1, v_Y1, v_X2 - v_X1 + 1, v_Y2 - v_Y1 + 1))
    End Sub

    Private Sub InvertColors(ByRef clrForeColor As Color, ByRef clrBackColor As Color)
        Dim clrBuff As Color
        clrBuff = clrBackColor
        clrBackColor = clrForeColor
        clrForeColor = clrBuff
    End Sub

    Private Function mp_GRect(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer) As RectangleGeometry
        Dim oReturn As New RectangleGeometry(New Rect(X, Y, Width, Height))
        Return oReturn
    End Function

    Private Function mp_GLine(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As LineGeometry
        If X1 <> X2 Then
            X2 = X2 + 1
        End If
        If Y1 <> Y2 Then
            Y2 = Y2 + 1
        End If
        Dim oReturn As New LineGeometry(New Point(X1, Y1), New Point(X2, Y2))
        Return oReturn
    End Function

    Private Function mp_GPoint(ByVal X1 As Integer, ByVal Y1 As Integer) As LineGeometry
        Dim oReturn As New LineGeometry(New Point(X1, Y1), New Point(X1 + 1, Y1 + 1))
        Return oReturn
    End Function

    Public Sub ResetFocusRectangle()
        mp_lSelectionRectangleIndex = -1
        mp_lSelectionLineIndex = -1
        mp_lFocusLeft = 0
        mp_lFocusTop = 0
        mp_lFocusRight = 0
        mp_lFocusBottom = 0
    End Sub

    Public Sub DrawReversibleFrameEx()
        DrawReversibleFrame(mp_lFocusLeft, mp_lFocusTop, mp_lFocusRight, mp_lFocusBottom)
    End Sub

    Public Sub DrawReversibleLine(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer)

        If mp_lSelectionLineIndex = -1 Then
            mp_oSelectionLine.X1 = v_X1
            mp_oSelectionLine.X2 = v_X2
            mp_oSelectionLine.Y1 = v_Y1
            mp_oSelectionLine.Y2 = v_Y2
            mp_oSelectionLine.Stroke = New SolidColorBrush(Colors.Blue)
            mp_oSelectionLine.StrokeThickness = 1
            mp_oSelectionLine.SnapsToDevicePixels = True
            mp_oSelectionLine.SetValue(RenderOptions.EdgeModeProperty, EdgeMode.Aliased)
            mp_oSelectionLine.IsHitTestVisible = False
            mp_oControl.f_Canvas.Children.Add(mp_oSelectionLine)
            mp_lSelectionLineIndex = mp_oControl.f_Canvas.Children.Count() - 1
        Else
            mp_oSelectionLine.X1 = v_X1
            mp_oSelectionLine.X2 = v_X2
            mp_oSelectionLine.Y1 = v_Y1
            mp_oSelectionLine.Y2 = v_Y2
        End If
    End Sub

    Public Sub EraseReversibleLines()
        If mp_lSelectionLineIndex > -1 Then
            mp_oControl.f_Canvas.Children.Remove(mp_oSelectionLine)
            mp_lSelectionLineIndex = -1
        End If
    End Sub

    Public Sub DrawReversibleFrame(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer)

        If mp_lSelectionRectangleIndex = -1 Then
            mp_oSelectionRectangle.Width = v_X2 - v_X1 + 1
            mp_oSelectionRectangle.Height = v_Y2 - v_Y1 + 1
            mp_oSelectionRectangle.SetValue(Canvas.LeftProperty, CDbl(v_X1))
            mp_oSelectionRectangle.SetValue(Canvas.TopProperty, CDbl(v_Y1))
            mp_oSelectionRectangle.Stroke = New SolidColorBrush(Colors.Blue)
            mp_oSelectionRectangle.StrokeThickness = 1
            mp_oSelectionRectangle.SnapsToDevicePixels = True
            mp_oSelectionRectangle.SetValue(RenderOptions.EdgeModeProperty, EdgeMode.Aliased)
            mp_oSelectionRectangle.IsHitTestVisible = False
            mp_oControl.f_Canvas.Children.Add(mp_oSelectionRectangle)
            mp_lSelectionRectangleIndex = mp_oControl.f_Canvas.Children.Count() - 1
        Else
            mp_oSelectionRectangle.Width = v_X2 - v_X1 + 1
            mp_oSelectionRectangle.Height = v_Y2 - v_Y1 + 1
            mp_oSelectionRectangle.SetValue(Canvas.LeftProperty, CDbl(v_X1))
            mp_oSelectionRectangle.SetValue(Canvas.TopProperty, CDbl(v_Y1))
        End If
    End Sub

    Public Sub EraseReversibleFrames()
        If mp_lSelectionRectangleIndex > -1 Then
            mp_oControl.f_Canvas.Children.Remove(mp_oSelectionRectangle)
            mp_lSelectionRectangleIndex = -1
        End If
    End Sub

    Public Sub StartPrintControl(ByVal DestHdc As Integer, ByVal XOrigin As Integer, ByVal YOrigin As Integer, ByVal XOriginExtents As Integer, ByVal YOriginExtents As Integer, ByVal MarginX As Integer, ByVal MarginY As Integer, ByVal DestScale As Integer, Optional ByVal FontRatio As Single = 1)

    End Sub

    Public Sub EndPrintControl()

    End Sub

    Public Function LineIntersection(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Boolean
        Return True
    End Function

    Friend Sub mp_DrawItemI(ByRef oTask As clsTask, ByVal sStyleIndex As String, ByVal Selected As Boolean, ByRef v_oStyle As clsStyle)
        Dim oStyle As clsStyle
        Dim oMilestoneStyle As clsMilestoneStyle
        If (v_oStyle Is Nothing) Then
            If mp_oControl.StrLib.StrIsNumeric(sStyleIndex) Then
                If mp_oControl.StrLib.StrCLng(sStyleIndex) < 0 Or mp_oControl.StrLib.StrCLng(sStyleIndex) > mp_oControl.Styles.Count Then
                    mp_oControl.mp_ErrorReport(SYS_ERRORS.STYLE_INVALID_INDEX, "Style object element not found when preparing to draw, invalid index", "mp_DrawItemI")
                    Return
                End If
            Else
                If mp_oControl.Styles.oCollection.m_bDoesKeyExist(sStyleIndex) = False Then
                    mp_oControl.mp_ErrorReport(SYS_ERRORS.STYLE_INVALID_KEY, "Style object element not found when preparing to draw, invalid key (""" & sStyleIndex & """)", "mp_DrawItemI")
                    Return
                End If
            End If
            oStyle = mp_oControl.Styles.FItem(sStyleIndex)
        Else
            oStyle = v_oStyle
        End If
        Select Case oStyle.Appearance
            Case E_STYLEAPPEARANCE.SA_FLAT
                oMilestoneStyle = oStyle.MilestoneStyle
                DrawFigure(mp_oControl.MathLib.GetXCoordinateFromDate(oTask.StartDate), oTask.Top, oTask.Bottom - oTask.Top, oTask.Bottom - oTask.Top, oMilestoneStyle.ShapeIndex, oMilestoneStyle.BorderColor, oMilestoneStyle.FillColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Case E_STYLEAPPEARANCE.SA_GRAPHICAL
                If oStyle.MilestoneStyle.Image Is Nothing Then

                Else
                    DrawImage(oStyle.MilestoneStyle.Image, oStyle.ImageAlignmentHorizontal, oStyle.ImageAlignmentVertical, oStyle.ImageXMargin, oStyle.ImageYMargin, oTask.Left, oTask.Right, oTask.Top, oTask.Bottom, oStyle.UseMask)
                End If
            Case Else
                oMilestoneStyle = oStyle.MilestoneStyle
                DrawFigure(mp_oControl.MathLib.GetXCoordinateFromDate(oTask.StartDate), oTask.Top, oTask.Bottom - oTask.Top, oTask.Bottom - oTask.Top, oMilestoneStyle.ShapeIndex, oMilestoneStyle.BorderColor, oMilestoneStyle.FillColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
        End Select
        mp_DrawItemText(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom, oTask.LeftTrim, oTask.RightTrim, oStyle, oTask.Text)
        If oStyle.SelectionRectangleStyle.Visible = True And Selected Then
            If oStyle.SelectionRectangleStyle.Mode = E_SELECTIONRECTANGLEMODE.SRM_DOTTED Then
                DrawFocusRectangle(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom)
            ElseIf oStyle.SelectionRectangleStyle.Mode = E_SELECTIONRECTANGLEMODE.SRM_COLOR Then
                DrawLine(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom, GRE_LINETYPE.LT_BORDER, oStyle.SelectionRectangleStyle.Color, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.SelectionRectangleStyle.BorderWidth)
            End If
        End If
    End Sub

    Friend Sub mp_DrawItemEx(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal sText As String, ByVal v_bIsSelected As Boolean, ByRef v_oImage As Image, ByVal v_lLeftTrim As Integer, ByVal v_lRightTrim As Integer, ByRef v_oStyle As clsStyle, ByVal clrBackColor As Color, ByVal clrForeColor As Color, ByVal clrStartGradientColor As Color, ByVal clrEndGradientColor As Color, ByVal clrHatchBackColor As Color, ByVal clrHatchForeColor As Color)
        Dim oStyle As clsStyle
        Dim oTaskStyle As clsTaskStyle
        If (v_oStyle Is Nothing) Then
            mp_oControl.mp_ErrorReport(SYS_ERRORS.STYLE_NULL, "Style object is null when preparing to draw.", "mp_DrawItemEx")
            Return
        Else
            oStyle = v_oStyle
        End If
        oTaskStyle = oStyle.TaskStyle
        Select Case oStyle.Appearance
            Case E_STYLEAPPEARANCE.SA_RAISED
                DrawEdge(v_lLeft, v_lTop, v_lRight, v_lBottom, clrBackColor, oStyle.ButtonStyle, GRE_EDGETYPE.ET_RAISED, True, v_oStyle)
            Case E_STYLEAPPEARANCE.SA_SUNKEN
                DrawEdge(v_lLeft, v_lTop, v_lRight, v_lBottom, clrBackColor, oStyle.ButtonStyle, GRE_EDGETYPE.ET_SUNKEN, True, v_oStyle)
            Case E_STYLEAPPEARANCE.SA_FLAT
                Dim lTop As Integer
                Dim lBottom As Integer
                lTop = v_lTop
                lBottom = v_lBottom
                Select Case oStyle.FillMode
                    Case GRE_FILLMODE.FM_COMPLETELYFILLED
                    Case GRE_FILLMODE.FM_UPPERHALFFILLED
                        lBottom = v_lTop + ((v_lBottom - v_lTop) / 2)
                    Case GRE_FILLMODE.FM_LOWERHALFFILLED
                        lTop = v_lBottom - ((v_lBottom - v_lTop) / 2)
                End Select
                If (oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID) Then
                    DrawLine(v_lLeft, lTop, v_lRight, lBottom, GRE_LINETYPE.LT_FILLED, clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                ElseIf (oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT) Then
                    GradientFill(v_lLeft, lTop, v_lRight, lBottom, clrStartGradientColor, clrEndGradientColor, oStyle.GradientFillMode)
                ElseIf (oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_PATTERN) Then
                    DrawPattern(v_lLeft, lTop, v_lRight, lBottom, clrBackColor, oStyle.Pattern, oStyle.PatternFactor)
                ElseIf (oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_HATCH) Then
                    HatchFill(v_lLeft, lTop, v_lRight, lBottom, clrHatchForeColor, clrHatchBackColor, oStyle.HatchStyle)
                End If
                If oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE Then
                    DrawLine(v_lLeft, lTop, v_lRight, lBottom, GRE_LINETYPE.LT_BORDER, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
                ElseIf oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM Then
                    If oStyle.CustomBorderStyle.Left = True Then
                        DrawLine(v_lLeft, lTop, v_lLeft, lBottom, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
                    End If
                    If oStyle.CustomBorderStyle.Top = True Then
                        DrawLine(v_lLeft, lTop, v_lRight, lTop, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
                    End If
                    If oStyle.CustomBorderStyle.Right = True Then
                        DrawLine(v_lRight, lTop, v_lRight, lBottom, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
                    End If
                    If oStyle.CustomBorderStyle.Bottom = True Then
                        DrawLine(v_lLeft, lBottom, v_lRight, lBottom, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
                    End If
                End If
                DrawFigure(v_lRight, v_lTop, v_lBottom - v_lTop, v_lBottom - v_lTop, oTaskStyle.EndShapeIndex, oTaskStyle.EndBorderColor, oTaskStyle.EndFillColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawFigure(v_lLeft, v_lTop, v_lBottom - v_lTop, v_lBottom - v_lTop, oTaskStyle.StartShapeIndex, oTaskStyle.StartBorderColor, oTaskStyle.StartFillColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Case E_STYLEAPPEARANCE.SA_CELL
                DrawLine(v_lLeft, v_lTop, v_lRight, v_lBottom, GRE_LINETYPE.LT_FILLED, clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(v_lLeft, v_lBottom, v_lRight, v_lBottom, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
            Case E_STYLEAPPEARANCE.SA_GRAPHICAL
                If oTaskStyle.MiddleImage Is Nothing Or oTaskStyle.StartImage Is Nothing Or oTaskStyle.EndImage Is Nothing Then

                Else
                    Dim lImageHeight As Integer
                    Dim lImageWidth As Integer
                    lImageHeight = oTaskStyle.MiddleImage.Height
                    lImageWidth = oTaskStyle.MiddleImage.Width
                    TileImageHorizontal(oTaskStyle.MiddleImage, v_lLeft, v_lTop, v_lRight, v_lBottom, oStyle.UseMask)
                    '// Exit if the start and end sections don't fit
                    If (v_lRight - v_lLeft) > (lImageWidth * 2) Then
                        '// Left Section
                        PaintImage(oTaskStyle.StartImage, v_lLeft, v_lTop, v_lLeft + lImageWidth, v_lTop + lImageHeight, 0, 0, oStyle.UseMask)
                        '// Right Section
                        PaintImage(oTaskStyle.EndImage, v_lRight - lImageWidth, v_lTop, v_lRight, v_lTop + lImageHeight, 0, 0, oStyle.UseMask)
                    End If
                End If
        End Select
        If Not (v_oImage Is Nothing) Then
            DrawImage(v_oImage, oStyle.ImageAlignmentHorizontal, oStyle.ImageAlignmentVertical, oStyle.ImageXMargin, oStyle.ImageYMargin, v_lLeft, v_lRight, v_lTop, v_lBottom, oStyle.UseMask)
        End If
        mp_DrawItemText(v_lLeft, v_lTop, v_lRight, v_lBottom, v_lLeftTrim, v_lRightTrim, oStyle, sText)
        If oStyle.SelectionRectangleStyle.Visible = True And v_bIsSelected Then
            mp_DrawSelectionRectangle(v_lLeft, v_lTop, v_lRight, v_lBottom, oStyle)
        End If
    End Sub

    Friend Sub mp_DrawSelectionRectangle(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal oStyle As clsStyle)
        If oStyle.SelectionRectangleStyle.Mode = E_SELECTIONRECTANGLEMODE.SRM_DOTTED Then
            DrawFocusRectangle(v_lLeft + oStyle.SelectionRectangleStyle.OffsetLeft, v_lTop + oStyle.SelectionRectangleStyle.OffsetTop, v_lRight - oStyle.SelectionRectangleStyle.OffsetRight, v_lBottom - oStyle.SelectionRectangleStyle.OffsetBottom)
        ElseIf oStyle.SelectionRectangleStyle.Mode = E_SELECTIONRECTANGLEMODE.SRM_COLOR Then
            DrawLine(v_lLeft + oStyle.SelectionRectangleStyle.OffsetLeft, v_lTop + oStyle.SelectionRectangleStyle.OffsetTop, v_lRight - oStyle.SelectionRectangleStyle.OffsetRight, v_lBottom - oStyle.SelectionRectangleStyle.OffsetBottom, GRE_LINETYPE.LT_BORDER, oStyle.SelectionRectangleStyle.Color, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.SelectionRectangleStyle.BorderWidth)
        End If
    End Sub

    Friend Sub mp_DrawItem(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal sStyleIndex As String, ByVal sText As String, ByVal v_bIsSelected As Boolean, ByRef v_oImage As Image, ByVal v_lLeftTrim As Integer, ByVal v_lRightTrim As Integer, ByRef v_oStyle As clsStyle)
        Dim oStyle As clsStyle
        If (v_oStyle Is Nothing) Then
            If mp_oControl.StrLib.StrIsNumeric(sStyleIndex) Then
                If mp_oControl.StrLib.StrCLng(sStyleIndex) < 0 Or mp_oControl.StrLib.StrCLng(sStyleIndex) > mp_oControl.Styles.Count Then
                    mp_oControl.mp_ErrorReport(SYS_ERRORS.STYLE_INVALID_INDEX, "Style object element not found when preparing to draw, invalid index", "mp_DrawItem")
                    Return
                End If
            Else
                If mp_oControl.Styles.oCollection.m_bDoesKeyExist(sStyleIndex) = False Then
                    mp_oControl.mp_ErrorReport(SYS_ERRORS.STYLE_INVALID_KEY, "Style object element not found when preparing to draw, invalid key (""" & sStyleIndex & """)", "mp_DrawItem")
                    Return
                End If
            End If
            oStyle = mp_oControl.Styles.FItem(sStyleIndex)
        Else
            oStyle = v_oStyle
        End If
        mp_DrawItemEx(v_lLeft, v_lTop, v_lRight, v_lBottom, sText, v_bIsSelected, v_oImage, v_lLeftTrim, v_lRightTrim, oStyle, oStyle.BackColor, oStyle.ForeColor, oStyle.StartGradientColor, oStyle.EndGradientColor, oStyle.HatchBackColor, oStyle.HatchForeColor)
    End Sub

    Private Sub mp_DrawItemText(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal v_lLeftTrim As Integer, ByVal v_lRightTrim As Integer, ByRef oStyle As clsStyle, ByVal sText As String)
        Dim lTextLeft As Integer
        Dim lTextRight As Integer
        Dim lTextTop As Integer
        Dim lTextBottom As Integer
        If oStyle.TextVisible = False Then
            Return
        End If
        If sText = "" Then
            Return
        End If
        Select Case oStyle.TextPlacement
            Case E_TEXTPLACEMENT.SCP_OBJECTEXTENTSPLACEMENT
                If (oStyle.DrawTextInVisibleArea = False) Then
                    lTextLeft = v_lLeft
                    lTextRight = v_lRight
                Else
                    lTextLeft = v_lLeftTrim
                    lTextRight = v_lRightTrim
                End If
                lTextTop = v_lTop
                lTextBottom = v_lBottom
                If oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT Then
                    lTextLeft = v_lLeft + oStyle.TextXMargin
                End If
                If oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT Then
                    lTextRight = v_lRight - oStyle.TextXMargin
                End If
                If oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP Then
                    lTextTop = v_lTop + oStyle.TextYMargin
                End If
                If oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM Then
                    lTextBottom = v_lBottom - oStyle.TextYMargin
                End If
                DrawAlignedText(lTextLeft, lTextTop, lTextRight, lTextBottom, sText, oStyle.TextAlignmentHorizontal, oStyle.TextAlignmentVertical, oStyle.ForeColor, oStyle.Font, oStyle.ClipText)
            Case E_TEXTPLACEMENT.SCP_OFFSETPLACEMENT
                DrawTextEx(v_lLeft + oStyle.TextFlags.OffsetLeft, v_lTop + oStyle.TextFlags.OffsetTop, v_lRight - oStyle.TextFlags.OffsetRight, v_lBottom - oStyle.TextFlags.OffsetBottom, sText, oStyle.TextFlags, oStyle.ForeColor, oStyle.Font, oStyle.ClipText)
            Case E_TEXTPLACEMENT.SCP_EXTERIORPLACEMENT
                If oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT Then
                    lTextLeft = v_lLeft - mp_oControl.mp_lStrWidth(sText, oStyle.Font) - oStyle.TextXMargin
                    lTextRight = v_lLeft - oStyle.TextXMargin + 1
                End If
                If oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT Then
                    lTextLeft = v_lRight + oStyle.TextXMargin
                    lTextRight = v_lRight + mp_oControl.mp_lStrWidth(sText, oStyle.Font) + oStyle.TextXMargin + 1
                End If
                If oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_CENTER Then
                    lTextLeft = v_lLeft
                    lTextRight = v_lRight + 1
                End If
                If oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP Then
                    lTextTop = v_lTop - mp_oControl.mp_lStrHeight(sText, oStyle.Font) - oStyle.TextYMargin
                    lTextBottom = v_lTop - oStyle.TextYMargin + 1
                End If
                If oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM Then
                    lTextTop = v_lBottom + oStyle.TextYMargin
                    lTextBottom = v_lBottom + mp_oControl.mp_lStrHeight(sText, oStyle.Font) + oStyle.TextYMargin + 1
                End If
                If oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_CENTER Then
                    lTextTop = v_lTop
                    lTextBottom = v_lBottom + 1
                End If
                DrawAlignedText(lTextLeft, lTextTop, lTextRight, lTextBottom, sText, GRE_HORIZONTALALIGNMENT.HAL_LEFT, GRE_VERTICALALIGNMENT.VAL_TOP, oStyle.ForeColor, oStyle.Font, oStyle.ClipText)
        End Select

    End Sub

    Friend Sub DrawScrollButton(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal width As Integer, ByVal height As Integer, ByVal button As E_SCROLLBUTTON, ByVal state As E_SCROLLBUTTONSTATE)
        Dim clrLightGray As Color = Color.FromRgb(192, 192, 192)
        Dim clrMediumGray As Color = Color.FromRgb(128, 128, 128)
        Dim clrDarkGray As Color = Color.FromRgb(64, 64, 64)
        Dim lOffset As Integer
        Dim lOffsetI As Integer
        Dim clrArrowColor As Color = Colors.Black
        DrawLine(X1 + 2, Y1 + 2, X1 + width - 3, Y1 + height - 3, GRE_LINETYPE.LT_FILLED, clrLightGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
        Select Case state
            Case E_SCROLLBUTTONSTATE.BS_NORMAL, E_SCROLLBUTTONSTATE.BS_INACTIVE
                DrawLine(X1, Y1, X1 + width - 2, Y1, GRE_LINETYPE.LT_NORMAL, clrLightGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1, Y1, X1, Y1 + height - 2, GRE_LINETYPE.LT_NORMAL, clrLightGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1, Y1 + height - 1, X1 + width - 1, Y1 + height - 1, GRE_LINETYPE.LT_NORMAL, clrDarkGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1 + width - 1, Y1, X1 + width - 1, Y1 + height - 1, GRE_LINETYPE.LT_NORMAL, clrDarkGray, GRE_LINEDRAWSTYLE.LDS_SOLID)

                DrawLine(X1 + 1, Y1 + 1, X1 + width - 3, Y1 + 1, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1 + 1, Y1 + 1, X1 + 1, Y1 + height - 3, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)

                DrawLine(X1 + 1, Y1 + height - 2, X1 + width - 2, Y1 + height - 2, GRE_LINETYPE.LT_NORMAL, clrMediumGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1 + width - 2, Y1 + 1, X1 + width - 2, Y1 + height - 2, GRE_LINETYPE.LT_NORMAL, clrMediumGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Case E_SCROLLBUTTONSTATE.BS_PUSHED
                DrawLine(X1, Y1, X1 + width - 2, Y1, GRE_LINETYPE.LT_NORMAL, clrMediumGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1, Y1, X1, Y1 + height - 2, GRE_LINETYPE.LT_NORMAL, clrMediumGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1, Y1 + height - 1, X1 + width - 1, Y1 + height - 1, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1 + width - 1, Y1, X1 + width - 1, Y1 + height - 1, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)

                DrawLine(X1 + 1, Y1 + 1, X1 + width - 3, Y1 + 1, GRE_LINETYPE.LT_NORMAL, clrDarkGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1 + 1, Y1 + 1, X1 + 1, Y1 + height - 3, GRE_LINETYPE.LT_NORMAL, clrDarkGray, GRE_LINEDRAWSTYLE.LDS_SOLID)

                DrawLine(X1 + 1, Y1 + height - 2, X1 + width - 2, Y1 + height - 2, GRE_LINETYPE.LT_NORMAL, clrLightGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1 + width - 2, Y1 + 1, X1 + width - 2, Y1 + height - 2, GRE_LINETYPE.LT_NORMAL, clrLightGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
        End Select
        If state = E_SCROLLBUTTONSTATE.BS_PUSHED Then
            lOffset = 1
        End If
        If state = E_SCROLLBUTTONSTATE.BS_INACTIVE Then
            clrArrowColor = clrMediumGray
            lOffsetI = 1
        End If
        Select Case button
            Case E_SCROLLBUTTON.SB_UP
                If state = E_SCROLLBUTTONSTATE.BS_INACTIVE Then
                    DrawPoint(X1 + 8 + lOffsetI, Y1 + 6 + lOffsetI, Colors.White)
                    DrawLine(X1 + 7 + lOffsetI, Y1 + 7 + lOffsetI, X1 + width - 8 + lOffsetI, Y1 + 7 + lOffsetI, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    DrawLine(X1 + 6 + lOffsetI, Y1 + 8 + lOffsetI, X1 + width - 7 + lOffsetI, Y1 + 8 + lOffsetI, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    DrawLine(X1 + 5 + lOffsetI, Y1 + 9 + lOffsetI, X1 + width - 6 + lOffsetI, Y1 + 9 + lOffsetI, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
                End If
                DrawPoint(X1 + 8 + lOffset, Y1 + 6 + lOffset, clrArrowColor)
                DrawLine(X1 + 7 + lOffset, Y1 + 7 + lOffset, X1 + width - 8 + lOffset, Y1 + 7 + lOffset, GRE_LINETYPE.LT_NORMAL, clrArrowColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1 + 6 + lOffset, Y1 + 8 + lOffset, X1 + width - 7 + lOffset, Y1 + 8 + lOffset, GRE_LINETYPE.LT_NORMAL, clrArrowColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1 + 5 + lOffset, Y1 + 9 + lOffset, X1 + width - 6 + lOffset, Y1 + 9 + lOffset, GRE_LINETYPE.LT_NORMAL, clrArrowColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Case E_SCROLLBUTTON.SB_DOWN
                If state = E_SCROLLBUTTONSTATE.BS_INACTIVE Then
                    DrawLine(X1 + 5 + lOffsetI, Y1 + 7 + lOffsetI, X1 + width - 6 + lOffsetI, Y1 + 7 + lOffsetI, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    DrawLine(X1 + 6 + lOffsetI, Y1 + 8 + lOffsetI, X1 + width - 7 + lOffsetI, Y1 + 8 + lOffsetI, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    DrawLine(X1 + 7 + lOffsetI, Y1 + 9 + lOffsetI, X1 + width - 8 + lOffsetI, Y1 + 9 + lOffsetI, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    DrawPoint(X1 + 8 + lOffsetI, Y1 + 10 + lOffsetI, Colors.White)
                End If
                DrawLine(X1 + 5 + lOffset, Y1 + 7 + lOffset, X1 + width - 6 + lOffset, Y1 + 7 + lOffset, GRE_LINETYPE.LT_NORMAL, clrArrowColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1 + 6 + lOffset, Y1 + 8 + lOffset, X1 + width - 7 + lOffset, Y1 + 8 + lOffset, GRE_LINETYPE.LT_NORMAL, clrArrowColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1 + 7 + lOffset, Y1 + 9 + lOffset, X1 + width - 8 + lOffset, Y1 + 9 + lOffset, GRE_LINETYPE.LT_NORMAL, clrArrowColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawPoint(X1 + 8 + lOffset, Y1 + 10 + lOffset, clrArrowColor)
            Case E_SCROLLBUTTON.SB_LEFT
                If state = E_SCROLLBUTTONSTATE.BS_INACTIVE Then
                    DrawPoint(X1 + 5 + lOffsetI, Y1 + 8 + lOffsetI, Colors.White)
                    DrawLine(X1 + 6 + lOffsetI, Y1 + 7 + lOffsetI, X1 + 6 + lOffsetI, Y1 + height - 8 + lOffsetI, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    DrawLine(X1 + 7 + lOffsetI, Y1 + 6 + lOffsetI, X1 + 7 + lOffsetI, Y1 + height - 7 + lOffsetI, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    DrawLine(X1 + 8 + lOffsetI, Y1 + 5 + lOffsetI, X1 + 8 + lOffsetI, Y1 + height - 6 + lOffsetI, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
                End If
                DrawPoint(X1 + 5 + lOffset, Y1 + 8 + lOffset, clrArrowColor)
                DrawLine(X1 + 6 + lOffset, Y1 + 7 + lOffset, X1 + 6 + lOffset, Y1 + height - 8 + lOffset, GRE_LINETYPE.LT_NORMAL, clrArrowColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1 + 7 + lOffset, Y1 + 6 + lOffset, X1 + 7 + lOffset, Y1 + height - 7 + lOffset, GRE_LINETYPE.LT_NORMAL, clrArrowColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1 + 8 + lOffset, Y1 + 5 + lOffset, X1 + 8 + lOffset, Y1 + height - 6 + lOffset, GRE_LINETYPE.LT_NORMAL, clrArrowColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Case E_SCROLLBUTTON.SB_RIGHT
                If state = E_SCROLLBUTTONSTATE.BS_INACTIVE Then
                    DrawLine(X1 + 7 + lOffsetI, Y1 + 5 + lOffsetI, X1 + 7 + lOffsetI, Y1 + height - 6 + lOffsetI, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    DrawLine(X1 + 8 + lOffsetI, Y1 + 6 + lOffsetI, X1 + 8 + lOffsetI, Y1 + height - 7 + lOffsetI, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    DrawLine(X1 + 9 + lOffsetI, Y1 + 7 + lOffsetI, X1 + 9 + lOffsetI, Y1 + height - 8 + lOffsetI, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    DrawPoint(X1 + 10 + lOffsetI, Y1 + 8 + lOffsetI, Colors.White)
                End If
                DrawLine(X1 + 7 + lOffset, Y1 + 5 + lOffset, X1 + 7 + lOffset, Y1 + height - 6 + lOffset, GRE_LINETYPE.LT_NORMAL, clrArrowColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1 + 8 + lOffset, Y1 + 6 + lOffset, X1 + 8 + lOffset, Y1 + height - 7 + lOffset, GRE_LINETYPE.LT_NORMAL, clrArrowColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(X1 + 9 + lOffset, Y1 + 7 + lOffset, X1 + 9 + lOffset, Y1 + height - 8 + lOffset, GRE_LINETYPE.LT_NORMAL, clrArrowColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawPoint(X1 + 10 + lOffset, Y1 + 8 + lOffset, clrArrowColor)
        End Select
    End Sub

    Public Sub DrawButton(ByVal oRect As Rect, ByVal state As E_SCROLLBUTTONSTATE)
        Dim clrLightGray As Color = Color.FromRgb(192, 192, 192)
        Dim clrMediumGray As Color = Color.FromRgb(128, 128, 128)
        Dim clrDarkGray As Color = Color.FromRgb(64, 64, 64)
        DrawLine(oRect.X + 1, oRect.Y + 1, oRect.X + oRect.Width - 3, oRect.Y + oRect.Height - 3, GRE_LINETYPE.LT_FILLED, clrLightGray, GRE_LINEDRAWSTYLE.LDS_SOLID)

        DrawLine(oRect.X, oRect.Y, oRect.X + oRect.Width - 2, oRect.Y, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
        DrawLine(oRect.X, oRect.Y, oRect.X, oRect.Y + oRect.Height - 2, GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
        DrawLine(oRect.X, oRect.Y + oRect.Height - 1, oRect.X + oRect.Width - 1, oRect.Y + oRect.Height - 1, GRE_LINETYPE.LT_NORMAL, clrDarkGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
        DrawLine(oRect.X + oRect.Width - 1, oRect.Y, oRect.X + oRect.Width - 1, oRect.Y + oRect.Height - 1, GRE_LINETYPE.LT_NORMAL, clrDarkGray, GRE_LINEDRAWSTYLE.LDS_SOLID)

        DrawLine(oRect.X + 1, oRect.Y + oRect.Height - 2, oRect.X + oRect.Width - 2, oRect.Y + oRect.Height - 2, GRE_LINETYPE.LT_NORMAL, clrMediumGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
        DrawLine(oRect.X + oRect.Width - 2, oRect.Y + 1, oRect.X + oRect.Width - 2, oRect.Y + oRect.Height - 2, GRE_LINETYPE.LT_NORMAL, clrMediumGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
    End Sub

    Friend Sub DrawPoint(ByVal X As Integer, ByVal Y As Integer, ByVal clrColor As Color)
        oGraphics.DrawLine(GetPen(clrColor), New Point(X, Y), New Point(X + 1, Y + 1))
    End Sub

    Friend Sub mp_DrawArrow(ByVal v_X As Integer, ByVal v_Y As Integer, ByVal v_ArrowDirection As GRE_ARROWDIRECTION, ByVal v_ArrowSize As Integer, ByVal v_lColor As Color)
        Dim i As Integer
        Select Case v_ArrowDirection
            Case GRE_ARROWDIRECTION.AWD_LEFT
                DrawPoint(v_X, v_Y, v_lColor)
                For i = 1 To v_ArrowSize
                    v_X = v_X + 1
                    DrawLine(v_X, v_Y - i, v_X, v_Y + i, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                Next
            Case GRE_ARROWDIRECTION.AWD_RIGHT
                DrawPoint(v_X, v_Y, v_lColor)
                For i = 1 To v_ArrowSize
                    v_X = v_X - 1
                    DrawLine(v_X, v_Y - i, v_X, v_Y + i, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                Next
            Case GRE_ARROWDIRECTION.AWD_UP
                DrawPoint(v_X, v_Y, v_lColor)
                For i = 1 To v_ArrowSize
                    v_Y = v_Y + 1
                    DrawLine(v_X - i, v_Y, v_X + i, v_Y, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                Next
            Case GRE_ARROWDIRECTION.AWD_DOWN
                DrawPoint(v_X, v_Y, v_lColor)
                For i = 1 To v_ArrowSize
                    v_Y = v_Y - 1
                    DrawLine(v_X - i, v_Y, v_X + i, v_Y, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                Next
        End Select
    End Sub

End Class
