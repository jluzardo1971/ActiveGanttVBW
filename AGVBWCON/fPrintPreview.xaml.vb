Imports System.Windows.Xps.Packaging
Imports System.IO
Imports System.IO.Packaging

Partial Public Class fPrintPreview

    Public oXPSDoc As XpsDocument
    Public oURI As Uri

    Private Sub fPrintPreview_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        Me.WindowState = Windows.WindowState.Maximized
        DocumentViewer1.Document = oXPSDoc.GetFixedDocumentSequence()
    End Sub

    Private Sub Resize()
        If Me.WindowState = Windows.WindowState.Normal Then
            DocumentViewer1.SetValue(Canvas.TopProperty, System.Convert.ToDouble(0))
            DocumentViewer1.SetValue(Canvas.LeftProperty, System.Convert.ToDouble(0))
            DocumentViewer1.Width = Me.Width - 8
            DocumentViewer1.Height = Me.Height - 30
        ElseIf Me.WindowState = Windows.WindowState.Maximized Then
            DocumentViewer1.SetValue(Canvas.TopProperty, System.Convert.ToDouble(0))
            DocumentViewer1.SetValue(Canvas.LeftProperty, System.Convert.ToDouble(0))
            DocumentViewer1.Width = SystemParameters.MaximizedPrimaryScreenWidth - 8
            DocumentViewer1.Height = SystemParameters.MaximizedPrimaryScreenHeight - 30
        End If
    End Sub

    Private Sub fPrintPreview_SizeChanged(ByVal sender As Object, ByVal e As System.Windows.SizeChangedEventArgs) Handles Me.SizeChanged
        Resize()
    End Sub

    Private Sub fPrintPreview_StateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.StateChanged
        Resize()
    End Sub

    Private Sub fPrintPreview_Unloaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Unloaded
        PackageStore.RemovePackage(oURI)
        oXPSDoc.Close()
    End Sub
End Class
