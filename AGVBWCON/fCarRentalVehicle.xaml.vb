Imports AGVBW
Imports System.Data

Partial Public Class fCarRentalVehicle

    Private mp_yDialogMode As PRG_DIALOGMODE
    Private mp_oParent As fCarRental
    Private mp_sRowID As String
    '// XML
    Private mp_otb_CR_ACRISS_Codes1 As DataTable
    Private mp_otb_CR_ACRISS_Codes2 As DataTable
    Private mp_otb_CR_ACRISS_Codes3 As DataTable
    Private mp_otb_CR_ACRISS_Codes4 As DataTable

    Friend Sub New(ByVal yDialogMode As PRG_DIALOGMODE, ByRef oParent As fCarRental, ByVal sRowID As String)
        MyBase.New()
        InitializeComponent()
        mp_yDialogMode = yDialogMode
        mp_oParent = oParent
        mp_sRowID = sRowID
    End Sub

    Private Sub fCarRentalVehicle_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            g_DST_ACCESS_FillComboBox(drpCarTypeID, "SELECT * FROM tb_CR_Car_Types", "lCarTypeID", "sDescription")
            g_DST_ACCESS_FillComboBox(drpACRISS1, "SELECT * FROM tb_CR_ACRISS_Codes WHERE [Position] = 1", "Letter", "Description")
            g_DST_ACCESS_FillComboBox(drpACRISS2, "SELECT * FROM tb_CR_ACRISS_Codes WHERE [Position] = 2", "Letter", "Description")
            g_DST_ACCESS_FillComboBox(drpACRISS3, "SELECT * FROM tb_CR_ACRISS_Codes WHERE [Position] = 3", "Letter", "Description")
            g_DST_ACCESS_FillComboBox(drpACRISS4, "SELECT * FROM tb_CR_ACRISS_Codes WHERE [Position] = 4", "Letter", "Description")
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            g_DST_XML_FillComboBox(drpCarTypeID, mp_oParent.mp_otb_CR_Car_Types.Tables(1), "lCarTypeID", "sDescription")
            mp_otb_CR_ACRISS_Codes1 = mp_oParent.mp_otb_CR_ACRISS_Codes.Tables(1).Copy
            mp_otb_CR_ACRISS_Codes2 = mp_oParent.mp_otb_CR_ACRISS_Codes.Tables(1).Copy
            mp_otb_CR_ACRISS_Codes3 = mp_oParent.mp_otb_CR_ACRISS_Codes.Tables(1).Copy
            mp_otb_CR_ACRISS_Codes4 = mp_oParent.mp_otb_CR_ACRISS_Codes.Tables(1).Copy
            g_DST_XML_FillComboBox(drpACRISS1, mp_otb_CR_ACRISS_Codes1, "Letter", "Description", "Position = 1")
            g_DST_XML_FillComboBox(drpACRISS2, mp_otb_CR_ACRISS_Codes2, "Letter", "Description", "Position = 2")
            g_DST_XML_FillComboBox(drpACRISS3, mp_otb_CR_ACRISS_Codes3, "Letter", "Description", "Position = 3")
            g_DST_XML_FillComboBox(drpACRISS4, mp_otb_CR_ACRISS_Codes4, "Letter", "Description", "Position = 4")
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            g_DST_NONE_FillComboBox(drpCarTypeID, mp_oParent.mp_otb_CR_Car_Types.Tables(0), "lCarTypeID", "sDescription")
            mp_otb_CR_ACRISS_Codes1 = mp_oParent.mp_otb_CR_ACRISS_Codes.Tables(0).Copy
            mp_otb_CR_ACRISS_Codes2 = mp_oParent.mp_otb_CR_ACRISS_Codes.Tables(0).Copy
            mp_otb_CR_ACRISS_Codes3 = mp_oParent.mp_otb_CR_ACRISS_Codes.Tables(0).Copy
            mp_otb_CR_ACRISS_Codes4 = mp_oParent.mp_otb_CR_ACRISS_Codes.Tables(0).Copy
            g_DST_NONE_FillComboBox(drpACRISS1, mp_otb_CR_ACRISS_Codes1, "Letter", "Description", "Position = 1")
            g_DST_NONE_FillComboBox(drpACRISS2, mp_otb_CR_ACRISS_Codes2, "Letter", "Description", "Position = 2")
            g_DST_NONE_FillComboBox(drpACRISS3, mp_otb_CR_ACRISS_Codes3, "Letter", "Description", "Position = 3")
            g_DST_NONE_FillComboBox(drpACRISS4, mp_otb_CR_ACRISS_Codes4, "Letter", "Description", "Position = 4")
        End If
        If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
            Me.Title = "Add New Vehicle"
            txtLicensePlates.Text = g_GenerateRandomLicense()
            drpCarTypeID.SelectedIndex = GetRnd(1, 48)
        ElseIf mp_yDialogMode = PRG_DIALOGMODE.DM_EDIT Then
            Dim oDataRow As DataRow = Nothing
            Me.Title = "Edit Vehicle"

            Dim sACRISSCode As String = ""

            If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                Dim oDB As clsDB = Nothing
                oDB = New clsDB()
                oDB.InitReader("SELECT * FROM tb_CR_Rows WHERE lRowID = " & mp_sRowID)
                oDB.Read(txtLicensePlates.Text, "sLicensePlates")
                oDB.Read(drpCarTypeID, "lCarTypeID")
                oDB.Read(txtNotes, "sNotes")
                oDB.Read(txtRate, "cRate")
                sACRISSCode = oDB.Read("sACRISSCode")
                oDB.CloseReader()
            ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                oDataRow = mp_oParent.mp_otb_CR_Rows.Tables(1).Rows.Find(mp_sRowID)
                txtLicensePlates.Text = DirectCast(oDataRow("sLicensePlates"), System.String)
                drpCarTypeID.SelectedValue = oDataRow("lCarTypeID")
                txtNotes.Text = DirectCast(oDataRow("sNotes"), System.String)
                txtRate.Text = CType(oDataRow("cRate"), System.String)
                sACRISSCode = DirectCast(oDataRow("sACRISSCode"), System.String)
            ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                oDataRow = mp_oParent.mp_otb_CR_Rows.Tables(0).Rows.Find(mp_sRowID)
                txtLicensePlates.Text = DirectCast(oDataRow("sLicensePlates"), System.String)
                drpCarTypeID.SelectedValue = oDataRow("lCarTypeID")
                txtNotes.Text = DirectCast(oDataRow("sNotes"), System.String)
                txtRate.Text = CType(oDataRow("cRate"), System.String)
                sACRISSCode = DirectCast(oDataRow("sACRISSCode"), System.String)
            End If
            UpdatePicture()
            UpdateACRISSCode(sACRISSCode)
        End If
    End Sub

    Private Sub UpdateACRISSCode(ByVal sACRISSCode As String)
        drpACRISS1.SelectedValue = sACRISSCode.Substring(0, 1)
        drpACRISS2.SelectedValue = sACRISSCode.Substring(1, 1)
        drpACRISS3.SelectedValue = sACRISSCode.Substring(2, 1)
        drpACRISS4.SelectedValue = sACRISSCode.Substring(3, 1)
        lblACRISS1.Content = sACRISSCode.Substring(0, 1)
        lblACRISS2.Content = sACRISSCode.Substring(1, 1)
        lblACRISS3.Content = sACRISSCode.Substring(2, 1)
        lblACRISS4.Content = sACRISSCode.Substring(3, 1)
    End Sub

    Private Sub UpdatePicture()
        If drpCarTypeID.SelectedItem.Item("sDescription") = "" Then
            Return
        End If
        picMake.Source = GetImageSource(g_GetAppLocation() & "\CarRental\Big\" & drpCarTypeID.SelectedItem.Item("sDescription") & ".jpg")
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdOK.Click
        Dim oRow As clsRow = Nothing
        Dim oDataRow As DataRow = Nothing
        If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            Dim oDB As clsDB = Nothing
            oDB = New clsDB()
            oDB.AddParameter("lDepth", 1, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("sLicensePlates", txtLicensePlates.Text, clsDB.ParamType.PT_STRING)
            oDB.AddParameter("lCarTypeID", g_GetComboBoxSelectedItem(drpCarTypeID, "lCarTypeID"), clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("sNotes", txtNotes.Text, clsDB.ParamType.PT_STRING)
            oDB.AddParameter("cRate", txtRate.Text, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("sACRISSCode", lblACRISS1.Content & lblACRISS2.Content & lblACRISS3.Content & lblACRISS4.Content, clsDB.ParamType.PT_STRING)
            If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
                oDB.AddParameter("lOrder", mp_oParent.ActiveGanttVBWCtl1.Rows.Count() + 1, clsDB.ParamType.PT_NUMERIC)
                mp_sRowID = "K" & oDB.ExecuteInsert("tb_CR_Rows")
                oRow = mp_oParent.ActiveGanttVBWCtl1.Rows.Add(mp_sRowID)
                oRow.Node.Depth = 1
                mp_oParent.ActiveGanttVBWCtl1.Rows.UpdateTree()
            ElseIf mp_yDialogMode = PRG_DIALOGMODE.DM_EDIT Then
                oDB.ExecuteUpdate("tb_CR_Rows", "lRowID = " & mp_sRowID)
                oRow = mp_oParent.ActiveGanttVBWCtl1.Rows.Item("K" & mp_sRowID)
            End If
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
                Dim lRowID As Integer = 0
                oDataRow = mp_oParent.mp_otb_CR_Rows.Tables(1).NewRow()
                lRowID = g_DST_XML_AutoIncrementValue(mp_oParent.mp_otb_CR_Rows, "lRowID")
                oDataRow("lRowID") = lRowID
                mp_sRowID = "K" & lRowID.ToString()
                oRow = mp_oParent.ActiveGanttVBWCtl1.Rows.Add(mp_sRowID)
                oRow.Node.Depth = 1
                mp_oParent.ActiveGanttVBWCtl1.Rows.UpdateTree()
                mp_oParent.mp_otb_CR_Rows.Tables(1).Rows.Add(oDataRow)
            ElseIf mp_yDialogMode = PRG_DIALOGMODE.DM_EDIT Then
                oDataRow = mp_oParent.mp_otb_CR_Rows.Tables(1).Rows.Find(mp_sRowID)
                oRow = mp_oParent.ActiveGanttVBWCtl1.Rows.Item("K" & mp_sRowID)
            End If
            oDataRow("lDepth") = 1
            oDataRow("sLicensePlates") = txtLicensePlates.Text
            oDataRow("lCarTypeID") = g_GetComboBoxSelectedItem(drpCarTypeID, "lCarTypeID")
            oDataRow("sNotes") = txtNotes.Text
            oDataRow("cRate") = txtRate.Text
            oDataRow("sACRISSCode") = lblACRISS1.Content & lblACRISS2.Content & lblACRISS3.Content & lblACRISS4.Content
            mp_oParent.mp_otb_CR_Rows.WriteXml(g_GetAppLocation() & "\CR_XML\tb_CR_Rows.xml")
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
                Dim lRowID As Integer = 0
                oDataRow = mp_oParent.mp_otb_CR_Rows.Tables(0).NewRow()
                lRowID = g_DST_NONE_AutoIncrementValue(mp_oParent.mp_otb_CR_Rows, "lRowID")
                oDataRow("lRowID") = lRowID
                mp_sRowID = "K" & lRowID.ToString()
                oRow = mp_oParent.ActiveGanttVBWCtl1.Rows.Add(mp_sRowID)
                oRow.Node.Depth = 1
                mp_oParent.ActiveGanttVBWCtl1.Rows.UpdateTree()
                mp_oParent.mp_otb_CR_Rows.Tables(0).Rows.Add(oDataRow)
            ElseIf mp_yDialogMode = PRG_DIALOGMODE.DM_EDIT Then
                oDataRow = mp_oParent.mp_otb_CR_Rows.Tables(0).Rows.Find(mp_sRowID)
                oRow = mp_oParent.ActiveGanttVBWCtl1.Rows.Item("K" & mp_sRowID)
            End If
            oDataRow("lDepth") = 1
            oDataRow("sLicensePlates") = txtLicensePlates.Text
            oDataRow("lCarTypeID") = g_GetComboBoxSelectedItem(drpCarTypeID, "lCarTypeID")
            oDataRow("sNotes") = txtNotes.Text
            oDataRow("cRate") = txtRate.Text
            oDataRow("sACRISSCode") = lblACRISS1.Content & lblACRISS2.Content & lblACRISS3.Content & lblACRISS4.Content
        End If

        oRow.Cells.Item("1").Text = txtLicensePlates.Text
        oRow.Cells.Item("2").Image = GetImage(g_GetAppLocation() & "\CarRental\Small\" & drpCarTypeID.SelectedItem.Item("sDescription") & ".jpg")
        oRow.Cells.Item("3").Text = g_GetComboBoxSelectedItem(drpCarTypeID, "sDescription") & vbCrLf & lblACRISS1.Content & lblACRISS2.Content & lblACRISS3.Content & lblACRISS4.Content & " - " & txtRate.Text & " USD"
        oRow.Tag = lblACRISS1.Content & lblACRISS2.Content & lblACRISS3.Content & lblACRISS4.Content & "|" & txtRate.Text & "|" & g_GetComboBoxSelectedItem(drpCarTypeID, "lCarTypeID")
        If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
            Dim l As Integer
            l = System.Math.Floor(mp_oParent.ActiveGanttVBWCtl1.CurrentViewObject.ClientArea.Height / 41)
            If ((mp_oParent.ActiveGanttVBWCtl1.Rows.Count - l + 2) > 0) Then
                mp_oParent.ActiveGanttVBWCtl1.VerticalScrollBar.Value = (mp_oParent.ActiveGanttVBWCtl1.Rows.Count - l + 2)
            End If
        End If
        mp_oParent.ActiveGanttVBWCtl1.Redraw()
        Me.Close()
    End Sub

    Private Sub drpCarTypeID_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles drpCarTypeID.SelectionChanged
        Dim sACRISSCode As String
        UpdatePicture()
        sACRISSCode = drpCarTypeID.SelectedItem.Item("sACRISSCode")
        UpdateACRISSCode(sACRISSCode)
        txtRate.Text = drpCarTypeID.SelectedItem.Item("cStdRate")
    End Sub

    Private Sub drpACRISS1_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles drpACRISS1.SelectionChanged
        lblACRISS1.Content = drpACRISS1.SelectedItem("Letter")
    End Sub

    Private Sub drpACRISS2_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles drpACRISS2.SelectionChanged
        lblACRISS2.Content = drpACRISS2.SelectedItem("Letter")
    End Sub

    Private Sub drpACRISS3_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles drpACRISS3.SelectionChanged
        lblACRISS3.Content = drpACRISS3.SelectedItem("Letter")
    End Sub

    Private Sub drpACRISS4_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles drpACRISS4.SelectionChanged
        lblACRISS4.Content = drpACRISS4.SelectedItem("Letter")
    End Sub

    Private Function GetImage(ByVal sImage As String) As Image
        Dim oDecoder As New JpegBitmapDecoder(GetURI(sImage), BitmapCreateOptions.None, BitmapCacheOption.None)
        Dim oBitmap As BitmapSource = oDecoder.Frames(0)
        Dim oReturn As New Image
        oReturn.Source = oBitmap
        Return oReturn
    End Function

    Private Function GetImageSource(ByVal sImage As String) As BitmapSource
        Dim oDecoder As New JpegBitmapDecoder(GetURI(sImage), BitmapCreateOptions.None, BitmapCacheOption.None)
        Dim oBitmap As BitmapSource = oDecoder.Frames(0)
        Return oBitmap
    End Function

    Private Function GetURI(ByVal sImage As String) As Uri
        Dim oURI As Uri = Nothing
        oURI = New Uri(sImage, UriKind.RelativeOrAbsolute)
        Return oURI
    End Function
End Class
