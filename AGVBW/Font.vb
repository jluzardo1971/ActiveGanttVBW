Public Class Font
    Private mp_sFamily As String
    Public Size As Single = 10
    Friend Italic As Boolean = False
    Friend Underline As Boolean = False
    Public FontStyle As FontStyle
    Public FontWeight As FontWeight
    Public VerticalAlignment As GRE_VERTICALALIGNMENT = GRE_VERTICALALIGNMENT.VAL_TOP
    Public HorizontalAlignment As GRE_HORIZONTALALIGNMENT = GRE_HORIZONTALALIGNMENT.HAL_LEFT

    Public Sub New(ByVal FamilyName As String, ByVal emSize As Single)
        mp_sFamily = FamilyName
        Size = emSize
    End Sub

    Public Sub New(ByVal FamilyName As String, ByVal emSize As Single, ByVal newStyle As FontWeight)
        mp_sFamily = FamilyName
        Size = emSize
        FontWeight = newStyle
    End Sub

    Public Function GetFontFamily() As FontFamily
        Dim oFontFamily As New FontFamily(mp_sFamily)
        Return oFontFamily
    End Function


    Public ReadOnly Property WPFFontSize() As Double
        Get
            Return (96 * Size / 72)
        End Get
    End Property

    Public Function Clone() As Font
        Return CType(Me.MemberwiseClone(), Font)
    End Function

    Public ReadOnly Property Name() As String
        Get
            Return mp_sFamily
        End Get
    End Property

    Public Property FamilyName() As String
        Get
            Return mp_sFamily
        End Get
        Set(ByVal value As String)
            mp_sFamily = value
        End Set
    End Property

    Friend ReadOnly Property Bold() As Boolean
        Get
            If FontWeight = FontWeights.Bold Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

End Class
