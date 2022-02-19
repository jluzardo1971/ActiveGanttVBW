Imports System.Windows.Input

Class fMain

    Dim mp_oParentNode As TreeViewItem



    Private Sub Window1_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        lblCopyright.Content = "Copyright ©2002-" & DateTime.Now.Year.ToString() & " The Source Code Store LLC. All Rights Reserved. All trademarks are property of their legal owner."

        AddTitleNode("AGEX", "ActiveGantt Examples:", 4, 5)
        AddNode("AGEX", "GanttCharts", "Gantt Charts:", 4, 5)

        AddNode("GanttCharts", "WBS", "Work Breakdown Structure (WBS) Project Management Examples:", 4, 5)
        AddNode("WBS", "WBSProject", "No data source (32bit and 64bit compatible)", 2, 2)
        AddNode("WBS", "WBSProjectXML", "XML data source (32bit and 64bit compatible)", 2, 2)
        AddNode("WBS", "WBSProjectAccess", "Microsoft Access data source (32bit compatible only)", 2, 2)

        AddNode("GanttCharts", "MSPI", "Microsoft Project Integration Examples (32bit and 64bit compatible):", 4, 5)
        AddNode("MSPI", "Project2003", "Demonstrates how ActiveGantt integrates with MS Project 2003 (using XML Files and the MSP2003 Integration Library)", 2, 2)
        AddNode("MSPI", "Project2007", "Demonstrates how ActiveGantt integrates with MS Project 2007 (using XML Files and the MSP2007 Integration Library)", 2, 2)
        AddNode("MSPI", "Project2010", "Demonstrates how ActiveGantt integrates with MS Project 2010 (using XML Files and the MSP2010 Integration Library)", 2, 2)

        AddNode("AGEX", "Schedules", "Schedules and Rosters:", 4, 5)

        AddNode("Schedules", "VRFC", "Vehicle Rental/Fleet Control Roster Examples:", 4, 5)
        AddNode("VRFC", "CarRental", "No data source (32bit and 64bit compatible)", 2, 2)
        AddNode("VRFC", "CarRentalXML", "XML data source (32bit and 64bit compatible)", 2, 2)
        AddNode("VRFC", "CarRentalAccess", "Microsoft Access data source (32bit compatible only)", 2, 2)

        AddNode("AGEX", "OTHER", "Other examples:", 4, 5)
        AddNode("OTHER", "FastLoad", "Fast Loading of Row and Task objects", 2, 2)
        AddNode("OTHER", "CustomDrawing", "Custom Drawing", 2, 2)
        AddNode("OTHER", "SortRows", "Sort Rows", 2, 2)
        AddNode("OTHER", "MillisecondInterval", "5 Millisecond Interval View", 2, 2)

        AddNode("OTHER", "TimeBlocks", "TimeBlocks and Duration Tasks:", 4, 5)
        AddNode("TimeBlocks", "RCT_DAY", "Daily Recurrent TimeBlocks", 2, 2)
        AddNode("TimeBlocks", "RCT_WEEK", "Weekly Recurrent TimeBlocks", 2, 2)
        AddNode("TimeBlocks", "RCT_MONTH", "Monthly Recurrent TimeBlocks", 2, 2)
        AddNode("TimeBlocks", "RCT_YEAR", "Yearly Recurrent TimeBlocks", 2, 2)
        AddNode("TimeBlocks", "DurationTasks", "Duration Tasks (can skip over non-working TimeBlock intervals)", 2, 2)

        AddTitleNode("HLP", "Help", 7, 7)
        AddNode("HLP", "GS_VBW", "How to create a simple WPF application using the ActiveGanttVBW component", 3, 3)
        AddNode("HLP", "LocalDocumentation", "ActiveGanttVBW Local Documentation", 7, 7)
        AddNode("HLP", "OnlineDocumentation", "ActiveGanttVBW Online Documentation", 6, 6)
        AddNode("HLP", "BugReport", "Submit a Bug Report", 3, 3)
        AddNode("HLP", "Request", "Request Further Explanations, Code Samples and Submit Technical Queries", 6, 6)

        AddTitleNode("SCS", "The Source Code Store LLC - Website (http://www.sourcecodestore.com/)", 3, 3)
        AddNode("SCS", "OnlineStore", "Online Store - Purchase ActiveGantt Online", 3, 3)
        AddNode("SCS", "ContactUs", "Contact Us (use this form for non technical queries only)", 3, 3)

        Dim oNode As TreeViewItem
        oNode = FindNode("OTHER")
        oNode.IsExpanded = False

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub AddTitleNode(ByVal sKey As String, ByVal sText As String, ByVal ImageIndex As Integer, ByVal SelectedImageIndex As Integer)
        Dim oNode As New TreeViewItem()
        oNode.Name = sKey
        oNode.Header = GetStackPanel(sText, ImageIndex)
        oNode.Tag = sKey
        oNode.IsExpanded = True
        TreeView1.Items.Add(oNode)
        mp_oParentNode = oNode
    End Sub

    Private Sub AddNode(ByVal sParentKey As String, ByVal sKey As String, ByVal sText As String, ByVal ImageIndex As Integer, ByVal SelectedImageIndex As Integer)
        Dim oNode As New TreeViewItem()
        Dim oParentNode As TreeViewItem
        oNode.Name = sKey
        oNode.Header = GetStackPanel(sText, ImageIndex)
        oNode.Tag = sKey
        oNode.IsExpanded = True
        oParentNode = FindNode(sParentKey)
        oParentNode.Items.Add(oNode)
    End Sub

    Private Function FindNode(ByVal sName As String) As TreeViewItem
        Dim i As Integer
        Dim oReturnTreeViewItem As TreeViewItem = Nothing
        For i = 0 To TreeView1.Items.Count - 1
            Dim oTreeViewItem As TreeViewItem = TreeView1.Items(i)
            oReturnTreeViewItem = FindNode_Intermediate(oTreeViewItem, sName)
            If Not (oReturnTreeViewItem Is Nothing) Then
                Return oReturnTreeViewItem
            End If
            oReturnTreeViewItem = FindNode_Final(oTreeViewItem, sName)
            If Not oReturnTreeViewItem Is Nothing Then
                Return oReturnTreeViewItem
            End If
        Next
        Return oReturnTreeViewItem
    End Function

    Private Function FindNode_Intermediate(ByRef oTreeViewItem As TreeViewItem, ByVal sName As String) As TreeViewItem
        Dim i As Integer
        Dim oReturnTreeViewItem As TreeViewItem = Nothing
        For i = 0 To oTreeViewItem.Items.Count - 1
            Dim oChildTreeViewItem As TreeViewItem = oTreeViewItem.Items(i)
            oReturnTreeViewItem = FindNode_Intermediate(oChildTreeViewItem, sName)
            If Not oReturnTreeViewItem Is Nothing Then
                Return oReturnTreeViewItem
            End If
        Next
        oReturnTreeViewItem = FindNode_Final(oTreeViewItem, sName)
        Return oReturnTreeViewItem
    End Function

    Private Function FindNode_Final(ByRef oTreeViewItem As TreeViewItem, ByVal sName As String) As TreeViewItem
        If oTreeViewItem.Name = sName Then
            Return oTreeViewItem
        End If
        Return Nothing
    End Function

    Private Function GetStackPanel(ByVal sText As String, ByVal ImageIndex As Integer) As StackPanel
        Dim oStackPanel As New StackPanel
        Dim oImage As New Image
        Dim oTextBlock As New TextBlock
        oImage.Source = GetImage(ImageIndex)
        oTextBlock.Text = " " & sText
        oStackPanel.Orientation = Orientation.Horizontal
        oStackPanel.Children.Add(oImage)
        oStackPanel.Children.Add(oTextBlock)
        Return oStackPanel
    End Function

    Private Function GetImage(ByVal ImageIndex As Integer) As BitmapSource
        Dim oDecoder As New GifBitmapDecoder(GetURI(ImageIndex), BitmapCreateOptions.None, BitmapCacheOption.None)
        Dim oBitmap As BitmapSource = oDecoder.Frames(0)
        Return oBitmap
    End Function

    Private Function GetURI(ByVal ImageIndex As Integer) As Uri
        Dim oURI As Uri = Nothing
        Select Case ImageIndex
            Case 4 'open folder
                oURI = New Uri("../Images/bfolderopen.gif", UriKind.RelativeOrAbsolute)
            Case 2 'ActiveGantt
                oURI = New Uri("../Images/AG.gif", UriKind.RelativeOrAbsolute)
            Case 3 'Inet
                oURI = New Uri("../Images/inet.gif", UriKind.RelativeOrAbsolute)
            Case 6
                oURI = New Uri("../Images/onlinedocumentation.gif", UriKind.RelativeOrAbsolute)
            Case 7
                oURI = New Uri("../Images/localCHMdocumentation.gif", UriKind.RelativeOrAbsolute)
        End Select
        Return oURI
    End Function

    Private Sub TreeView1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles TreeView1.MouseDoubleClick
        If TreeView1.SelectedItem Is Nothing Then
            Return
        End If
        Dim sSelectedTag As String = TreeView1.SelectedItem.Tag
        Select Case sSelectedTag
            Case "WBSProject"
                Dim oForm As New fWBSProject(E_DATASOURCETYPE.DST_NONE)
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "WBSProjectXML"
                Dim oForm As New fWBSProject(E_DATASOURCETYPE.DST_XML)
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "WBSProjectAccess"
                Dim oForm As New fWBSProject(E_DATASOURCETYPE.DST_ACCESS)
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "CarRental"
                Dim oForm As New fCarRental(E_DATASOURCETYPE.DST_NONE)
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "CarRentalXML"
                Dim oForm As New fCarRental(E_DATASOURCETYPE.DST_XML)
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "CarRentalAccess"
                Dim oForm As New fCarRental(E_DATASOURCETYPE.DST_ACCESS)
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "Project2003"
                Dim oForm As New fMSProject11()
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "Project2007"
                Dim oForm As New fMSProject12()
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "Project2010"
                Dim oForm As New fMSProject14()
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "FastLoad"
                Dim oForm As New fFastLoading()
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "CustomDrawing"
                Dim oForm As New fCustomDrawing()
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "SortRows"
                Dim oForm As New fSortRows()
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "MillisecondInterval"
                Dim oForm As New fMillisecondInterval()
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "DurationTasks"
                Dim oForm As New fDurationTasks
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "RCT_DAY"
                Dim oForm As New fRCT_DAY
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "RCT_WEEK"
                Dim oForm As New fRCT_WEEK
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "RCT_MONTH"
                Dim oForm As New fRCT_MONTH
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "RCT_YEAR"
                Dim oForm As New fRCT_YEAR
                Me.Cursor = Cursors.Wait
                oForm.ShowDialog()
                Me.Cursor = Cursors.Arrow
            Case "SCS"
                Me.Cursor = Cursors.Wait
                System.Diagnostics.Process.Start("http://www.sourcecodestore.com/")
                Me.Cursor = Cursors.Arrow
            Case "GS_VBW"
                Me.Cursor = Cursors.Wait
                System.Diagnostics.Process.Start("http://www.sourcecodestore.com/Article.aspx?ID=20#Create")
                Me.Cursor = Cursors.Arrow
            Case "OnlineStore"
                Me.Cursor = Cursors.Wait
                System.Diagnostics.Process.Start("http://www.sourcecodestore.com/OnlineStore/")
                Me.Cursor = Cursors.Arrow
            Case "OnlineDocumentation"
                Me.Cursor = Cursors.Wait
                System.Diagnostics.Process.Start("http://www.sourcecodestore.com/Documentation/DOCFrameset.aspx?PN=AG&PL=VBW")
                Me.Cursor = Cursors.Arrow
            Case "LocalDocumentation"
                Me.Cursor = Cursors.Wait
                System.Diagnostics.Process.Start(g_GetAppLocation() & "\AGVBW.chm")
                Me.Cursor = Cursors.Arrow
            Case "BugReport"
                Me.Cursor = Cursors.Wait
                System.Diagnostics.Process.Start("http://www.sourcecodestore.com/Support/Report.aspx?T=1")
                Me.Cursor = Cursors.Arrow
            Case "Request"
                Me.Cursor = Cursors.Wait
                System.Diagnostics.Process.Start("http://www.sourcecodestore.com/Support/Report.aspx?T=2")
                Me.Cursor = System.Windows.Input.Cursors.Arrow
            Case "ContactUs"
                Me.Cursor = Cursors.Wait
                System.Diagnostics.Process.Start("http://www.sourcecodestore.com/contactus.aspx")
                Me.Cursor = Cursors.Arrow
        End Select
    End Sub
End Class
