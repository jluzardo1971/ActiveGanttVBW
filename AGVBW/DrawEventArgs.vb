Option Explicit On 

Imports System
Imports System.ComponentModel
Imports System.Reflection

Public Class DrawEventArgs
    Inherits System.EventArgs

    Public EventTarget As E_EVENTTARGET
    Public CustomDraw As Boolean
    Public ObjectIndex As Integer
    Public ParentObjectIndex As Integer
    Public Graphics As DrawingContext

    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        EventTarget = Nothing
        CustomDraw = Nothing
        ObjectIndex = Nothing
        ParentObjectIndex = Nothing
        Graphics = Nothing
    End Sub
End Class

