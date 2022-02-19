Option Explicit On 

Imports System
Imports System.ComponentModel
Imports System.Reflection


Public Class ToolTipEventArgs
    Inherits System.EventArgs

    Public InitialRowIndex As Integer
    Public FinalRowIndex As Integer
    Public TaskIndex As Integer
    Public MilestoneIndex As Integer
    Public PercentageIndex As Integer
    Public RowIndex As Integer
    Public CellIndex As Integer
    Public ColumnIndex As Integer
    Public InitialStartDate As AGVBW.DateTime
    Public InitialEndDate As AGVBW.DateTime
    Public StartDate As AGVBW.DateTime
    Public EndDate As AGVBW.DateTime
    Public XStart As Integer
    Public XEnd As Integer
    Public Operation As E_OPERATION
    Public EventTarget As E_EVENTTARGET
    Public TaskPosition As String
    Public PredecessorPosition As String
    Public X As Integer
    Public Y As Integer
    Public X1 As Integer
    Public Y1 As Integer
    Public X2 As Integer
    Public Y2 As Integer
    Public CustomDraw As Boolean
    Public Graphics As Canvas
    Public ToolTipType As E_TOOLTIPTYPE

    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        InitialRowIndex = Nothing
        FinalRowIndex = Nothing
        RowIndex = Nothing
        TaskIndex = Nothing
        MilestoneIndex = Nothing
        PercentageIndex = Nothing
        CellIndex = Nothing
        ColumnIndex = Nothing
        StartDate = New AGVBW.DateTime()
        EndDate = New AGVBW.DateTime()
        InitialStartDate = New AGVBW.DateTime()
        InitialEndDate = New AGVBW.DateTime()
        XStart = Nothing
        XEnd = Nothing
        X = Nothing
        Y = Nothing
        X1 = Nothing
        Y1 = Nothing
        X2 = Nothing
        Y2 = Nothing
        Operation = Nothing
        EventTarget = Nothing
        ToolTipType = Nothing
    End Sub
End Class

