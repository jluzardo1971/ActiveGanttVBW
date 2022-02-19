Partial Friend Class fAbout

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdOK.Click
        Me.Close()
    End Sub

    Private Sub lblRegister_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles lblRegister.MouseUp
        System.Diagnostics.Process.Start(lblRegister.Tag)
    End Sub

    Private Sub Window1_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        Dim ai As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly
        Dim oStream As System.IO.Stream = ai.GetManifestResourceStream("AGVBW.AG.bmp")
        Dim oDecoder As BmpBitmapDecoder = New BmpBitmapDecoder(oStream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.None)
        lblTitle1.Content = "ActiveGantt Scheduler Control, Version " & ai.GetName.Version().ToString()
        lblCopyright.Content = "Copyright © 2002-" & System.DateTime.Now.Year.ToString() & ",  The Source Code Store LLC"
        lblURL.Content = "http://www.sourcecodestore.com"
        lblTechnicalSupport.Content = "Technical Support Page"
        lblSales.Content = "sales@sourcecodestore.com"
        lblRegister.Tag = "http://www.sourcecodestore.com/OnlineStore/default.aspx"
        picIcon.Source = oDecoder.Frames(0)
    End Sub

    Private Sub lblURL_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles lblURL.MouseUp
        System.Diagnostics.Process.Start(lblURL.Content)
    End Sub

    Private Sub lblTechnicalSupport_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles lblTechnicalSupport.MouseUp
        System.Diagnostics.Process.Start("http://www.sourcecodestore.com/Support/default.aspx")
    End Sub

    Private Sub lblSales_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles lblSales.MouseUp
        System.Diagnostics.Process.Start("mailto:" & lblSales.Content)
    End Sub
End Class
