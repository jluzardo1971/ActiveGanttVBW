Imports System.Data.OleDb
Imports AGVBW
Imports System.Data

Partial Public Class fCarRentalReservation

    Private mp_yDialogMode As PRG_DIALOGMODE
    Private mp_oParent As fCarRental
    Private mp_sTaskID As String

    Friend Sub New(ByVal yDialogMode As PRG_DIALOGMODE, ByRef oParent As fCarRental, ByVal sTaskID As String)
        MyBase.New()
        InitializeComponent()
        mp_yDialogMode = yDialogMode
        mp_oParent = oParent
        mp_sTaskID = sTaskID
    End Sub

    Private Sub fCarRentalReservation_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        Dim oTask As clsTask = GetCurrentTask()
        Dim sRowTag() As String
        Dim oDataRow As DataRow = Nothing
        If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
            Dim sCityName As String = ""
            Dim sStateName As String = ""
            Dim lID As Integer = 0
            If mp_oParent.Mode = fCarRental.HPE_ADDMODE.AM_RESERVATION Then
                Me.Title = "Add Reservation"
                lblMode.Content = "Add Reservation"
                lblMode.Background = New SolidColorBrush(Color.FromRgb(153, 170, 194))
            Else
                Me.Title = "Add Rental"
                lblMode.Content = "Add Rental"
                lblMode.Background = New SolidColorBrush(Color.FromRgb(162, 78, 50))
            End If
            g_GenerateRandomCity(sCityName, sStateName, lID, mp_oParent.mp_yDataSourceType)
            txtCity.Text = sCityName
            txtName.Text = g_GenerateRandomName(False, mp_oParent.mp_yDataSourceType)
            txtState.Text = sStateName
            txtPhone.Text = g_GenerateRandomPhone("")
            txtMobile.Text = g_GenerateRandomPhone(txtPhone.Text.Substring(0, 5))
            txtAddress.Text = g_GenerateRandomAddress(mp_oParent.mp_yDataSourceType)
            txtZIP.Text = g_GenerateRandomZIP()
            SetStartDate(oTask.StartDate)
            SetEndDate(oTask.EndDate)


            If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
                    Dim oReader As OleDbDataReader = Nothing
                    oReader = g_DST_ACCESS_ReturnReader("SELECT * FROM tb_CR_Taxes_Surcharges_Options", oConn)
                    While oReader.Read = True
                        If DirectCast(oReader.Item("sID"), System.String) = "GPS" Then
                            chkGPS.Content = DirectCast(oReader.Item("sDescription"), System.String)
                            chkGPS.Tag = oReader.Item("sRate")
                        End If
                        If DirectCast(oReader.Item("sID"), System.String) = "LDW" Then
                            chkLDW.Content = DirectCast(oReader.Item("sDescription"), System.String)
                            chkLDW.Tag = oReader.Item("sRate")
                        End If
                        If DirectCast(oReader.Item("sID"), System.String) = "PAI" Then
                            chkPAI.Content = DirectCast(oReader.Item("sDescription"), System.String)
                            chkPAI.Tag = oReader.Item("sRate")
                        End If
                        If DirectCast(oReader.Item("sID"), System.String) = "PEP" Then
                            chkPEP.Content = DirectCast(oReader.Item("sDescription"), System.String)
                            chkPEP.Tag = oReader.Item("sRate")
                        End If
                        If DirectCast(oReader.Item("sID"), System.String) = "ALI" Then
                            chkALI.Content = DirectCast(oReader.Item("sDescription"), System.String)
                            chkALI.Tag = oReader.Item("sRate")
                        End If
                        If DirectCast(oReader.Item("sID"), System.String) = "ERF" Then
                            txtERF.Tag = oReader.Item("sRate")
                        End If
                        If DirectCast(oReader.Item("sID"), System.String) = "CRF" Then
                            txtCRF.Tag = oReader.Item("sRate")
                        End If
                        If DirectCast(oReader.Item("sID"), System.String) = "RCFC" Then
                            txtRCFC.Tag = oReader.Item("sRate")
                        End If
                        If DirectCast(oReader.Item("sID"), System.String) = "WTB" Then
                            txtWTB.Tag = oReader.Item("sRate")
                        End If
                        If DirectCast(oReader.Item("sID"), System.String) = "VLF" Then
                            txtVLF.Tag = oReader.Item("sRate")
                        End If
                    End While
                    oReader.Close()
                End Using
            ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                For Each oDataRow1 As DataRow In mp_oParent.mp_otb_CR_Taxes_Surcharges_Options.Tables(1).Rows()
                    If DirectCast(oDataRow1.Item("sID"), System.String) = "GPS" Then
                        chkGPS.Content = DirectCast(oDataRow1.Item("sDescription"), System.String)
                        chkGPS.Tag = oDataRow1.Item("sRate")
                    End If
                    If DirectCast(oDataRow1.Item("sID"), System.String) = "LDW" Then
                        chkLDW.Content = DirectCast(oDataRow1.Item("sDescription"), System.String)
                        chkLDW.Tag = oDataRow1.Item("sRate")
                    End If
                    If DirectCast(oDataRow1.Item("sID"), System.String) = "PAI" Then
                        chkPAI.Content = DirectCast(oDataRow1.Item("sDescription"), System.String)
                        chkPAI.Tag = oDataRow1.Item("sRate")
                    End If
                    If DirectCast(oDataRow1.Item("sID"), System.String) = "PEP" Then
                        chkPEP.Content = DirectCast(oDataRow1.Item("sDescription"), System.String)
                        chkPEP.Tag = oDataRow1.Item("sRate")
                    End If
                    If DirectCast(oDataRow1.Item("sID"), System.String) = "ALI" Then
                        chkALI.Content = DirectCast(oDataRow1.Item("sDescription"), System.String)
                        chkALI.Tag = oDataRow1.Item("sRate")
                    End If
                    If DirectCast(oDataRow1.Item("sID"), System.String) = "ERF" Then
                        txtERF.Tag = oDataRow1.Item("sRate")
                    End If
                    If DirectCast(oDataRow1.Item("sID"), System.String) = "CRF" Then
                        txtCRF.Tag = oDataRow1.Item("sRate")
                    End If
                    If DirectCast(oDataRow1.Item("sID"), System.String) = "RCFC" Then
                        txtRCFC.Tag = oDataRow1.Item("sRate")
                    End If
                    If DirectCast(oDataRow1.Item("sID"), System.String) = "WTB" Then
                        txtWTB.Tag = oDataRow1.Item("sRate")
                    End If
                    If DirectCast(oDataRow1.Item("sID"), System.String) = "VLF" Then
                        txtVLF.Tag = oDataRow1.Item("sRate")
                    End If
                Next
            ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                For Each oDataRow2 As DataRow In mp_oParent.mp_otb_CR_Taxes_Surcharges_Options.Tables(0).Rows()
                    If DirectCast(oDataRow2.Item("sID"), System.String) = "GPS" Then
                        chkGPS.Content = DirectCast(oDataRow2.Item("sDescription"), System.String)
                        chkGPS.Tag = oDataRow2.Item("sRate")
                    End If
                    If DirectCast(oDataRow2.Item("sID"), System.String) = "LDW" Then
                        chkLDW.Content = DirectCast(oDataRow2.Item("sDescription"), System.String)
                        chkLDW.Tag = oDataRow2.Item("sRate")
                    End If
                    If DirectCast(oDataRow2.Item("sID"), System.String) = "PAI" Then
                        chkPAI.Content = DirectCast(oDataRow2.Item("sDescription"), System.String)
                        chkPAI.Tag = oDataRow2.Item("sRate")
                    End If
                    If DirectCast(oDataRow2.Item("sID"), System.String) = "PEP" Then
                        chkPEP.Content = DirectCast(oDataRow2.Item("sDescription"), System.String)
                        chkPEP.Tag = oDataRow2.Item("sRate")
                    End If
                    If DirectCast(oDataRow2.Item("sID"), System.String) = "ALI" Then
                        chkALI.Content = DirectCast(oDataRow2.Item("sDescription"), System.String)
                        chkALI.Tag = oDataRow2.Item("sRate")
                    End If
                    If DirectCast(oDataRow2.Item("sID"), System.String) = "ERF" Then
                        txtERF.Tag = oDataRow2.Item("sRate")
                    End If
                    If DirectCast(oDataRow2.Item("sID"), System.String) = "CRF" Then
                        txtCRF.Tag = oDataRow2.Item("sRate")
                    End If
                    If DirectCast(oDataRow2.Item("sID"), System.String) = "RCFC" Then
                        txtRCFC.Tag = oDataRow2.Item("sRate")
                    End If
                    If DirectCast(oDataRow2.Item("sID"), System.String) = "WTB" Then
                        txtWTB.Tag = oDataRow2.Item("sRate")
                    End If
                    If DirectCast(oDataRow2.Item("sID"), System.String) = "VLF" Then
                        txtVLF.Tag = oDataRow2.Item("sRate")
                    End If
                Next
            End If
        ElseIf mp_yDialogMode = PRG_DIALOGMODE.DM_EDIT Then
            If CType(oTask.Tag, System.Int32) = 0 Then
                Me.Title = "Edit Reservation"
                lblMode.Content = "Edit Reservation"
                lblMode.Background = New SolidColorBrush(Color.FromRgb(153, 170, 194))
            Else
                Me.Title = "Edit Rental"
                lblMode.Content = "Edit Rental"
                lblMode.Background = New SolidColorBrush(Color.FromRgb(162, 78, 50))
            End If

            If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                Dim oDB As clsDB = Nothing
                oDB = New clsDB()
                oDB.InitReader("SELECT * FROM tb_CR_Rentals WHERE lTaskID = " & mp_sTaskID)
                oDB.Read(txtCity, "sCity")
                oDB.Read(txtName, "sName")
                oDB.Read(txtState, "sState")
                oDB.Read(txtPhone, "sPhone")
                oDB.Read(txtMobile, "sMobile")
                oDB.Read(txtAddress, "sAddress")
                oDB.Read(txtZIP, "sZIP")
                SetStartDate(FromDate(System.Convert.ToDateTime(oDB.Read("dtPickUp"))))
                SetEndDate(FromDate(System.Convert.ToDateTime(oDB.Read("dtReturn"))))
                oDB.Read(txtRate, "cRate")
                chkGPS.Tag = CType(oDB.Read("cGPS"), System.Double)
                chkLDW.Tag = CType(oDB.Read("cLDW"), System.Double)
                chkPAI.Tag = CType(oDB.Read("cPAI"), System.Double)
                chkPEP.Tag = CType(oDB.Read("cPEP"), System.Double)
                chkALI.Tag = CType(oDB.Read("cALI"), System.Double)
                txtERF.Tag = CType(oDB.Read("cERF"), System.Double)
                txtCRF.Tag = CType(oDB.Read("dCRF"), System.Double)
                txtRCFC.Tag = CType(oDB.Read("cRCFC"), System.Double)
                txtWTB.Tag = CType(oDB.Read("cWTB"), System.Double)
                txtVLF.Tag = CType(oDB.Read("cVLF"), System.Double)
                lblTax.Tag = CType(oDB.Read("dTax"), System.Double)
                txtEstimatedTotal.Tag = CType(oDB.Read("cEstimatedTotal"), System.Double)
                oDB.Read(chkGPS, "bGPS")
                oDB.Read(chkFSO, "bFSO")
                oDB.Read(chkLDW, "bLDW")
                oDB.Read(chkPAI, "bPAI")
                oDB.Read(chkPEP, "bPEP")
                oDB.Read(chkALI, "bALI")
                oDB.CloseReader()
            ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                oDataRow = mp_oParent.mp_otb_CR_Rentals.Tables(1).Rows.Find(mp_sTaskID)
                txtCity.Text = DirectCast(oDataRow("sCity"), System.String)
                txtName.Text = DirectCast(oDataRow("sName"), System.String)
                txtState.Text = DirectCast(oDataRow("sState"), System.String)
                txtPhone.Text = DirectCast(oDataRow("sPhone"), System.String)
                txtMobile.Text = DirectCast(oDataRow("sMobile"), System.String)
                txtAddress.Text = DirectCast(oDataRow("sAddress"), System.String)
                txtZIP.Text = DirectCast(oDataRow("sZIP"), System.String)
                SetStartDate(FromDate(oDataRow("dtPickUp")))
                SetEndDate(FromDate(oDataRow("dtReturn")))
                txtRate.Text = CType(oDataRow("cRate"), System.String)
                chkGPS.Tag = DirectCast(oDataRow("cGPS"), System.Double)
                chkLDW.Tag = DirectCast(oDataRow("cLDW"), System.Double)
                chkPAI.Tag = DirectCast(oDataRow("cPAI"), System.Double)
                chkPEP.Tag = DirectCast(oDataRow("cPEP"), System.Double)
                chkALI.Tag = DirectCast(oDataRow("cALI"), System.Double)
                txtERF.Tag = DirectCast(oDataRow("cERF"), System.Double)
                txtCRF.Tag = DirectCast(oDataRow("dCRF"), System.Double)
                txtRCFC.Tag = DirectCast(oDataRow("cRCFC"), System.Double)
                txtWTB.Tag = DirectCast(oDataRow("cWTB"), System.Double)
                txtVLF.Tag = DirectCast(oDataRow("cVLF"), System.Double)
                lblTax.Tag = DirectCast(oDataRow("dTax"), System.Double)
                txtEstimatedTotal.Tag = DirectCast(oDataRow("cEstimatedTotal"), System.Double)
                chkGPS.IsChecked = DirectCast(oDataRow("bGPS"), System.Boolean)
                chkFSO.IsChecked = DirectCast(oDataRow("bFSO"), System.Boolean)
                chkLDW.IsChecked = DirectCast(oDataRow("bLDW"), System.Boolean)
                chkPAI.IsChecked = DirectCast(oDataRow("bPAI"), System.Boolean)
                chkPEP.IsChecked = DirectCast(oDataRow("bPEP"), System.Boolean)
                chkALI.IsChecked = DirectCast(oDataRow("bALI"), System.Boolean)
            ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                oDataRow = mp_oParent.mp_otb_CR_Rentals.Tables(0).Rows.Find(mp_sTaskID)
                txtCity.Text = DirectCast(oDataRow("sCity"), System.String)
                txtName.Text = DirectCast(oDataRow("sName"), System.String)
                txtState.Text = DirectCast(oDataRow("sState"), System.String)
                txtPhone.Text = DirectCast(oDataRow("sPhone"), System.String)
                txtMobile.Text = DirectCast(oDataRow("sMobile"), System.String)
                txtAddress.Text = DirectCast(oDataRow("sAddress"), System.String)
                txtZIP.Text = DirectCast(oDataRow("sZIP"), System.String)
                SetStartDate(FromDate(oDataRow("dtPickUp")))
                SetEndDate(FromDate(oDataRow("dtReturn")))
                txtRate.Text = CType(oDataRow("cRate"), System.String)
                chkGPS.Tag = DirectCast(oDataRow("cGPS"), System.Double)
                chkLDW.Tag = DirectCast(oDataRow("cLDW"), System.Double)
                chkPAI.Tag = DirectCast(oDataRow("cPAI"), System.Double)
                chkPEP.Tag = DirectCast(oDataRow("cPEP"), System.Double)
                chkALI.Tag = DirectCast(oDataRow("cALI"), System.Double)
                txtERF.Tag = DirectCast(oDataRow("cERF"), System.Double)
                txtCRF.Tag = DirectCast(oDataRow("dCRF"), System.Double)
                txtRCFC.Tag = DirectCast(oDataRow("cRCFC"), System.Double)
                txtWTB.Tag = DirectCast(oDataRow("cWTB"), System.Double)
                txtVLF.Tag = DirectCast(oDataRow("cVLF"), System.Double)
                lblTax.Tag = DirectCast(oDataRow("dTax"), System.Double)
                txtEstimatedTotal.Tag = DirectCast(oDataRow("cEstimatedTotal"), System.Double)
                chkGPS.IsChecked = DirectCast(oDataRow("bGPS"), System.Boolean)
                chkFSO.IsChecked = DirectCast(oDataRow("bFSO"), System.Boolean)
                chkLDW.IsChecked = DirectCast(oDataRow("bLDW"), System.Boolean)
                chkPAI.IsChecked = DirectCast(oDataRow("bPAI"), System.Boolean)
                chkPEP.IsChecked = DirectCast(oDataRow("bPEP"), System.Boolean)
                chkALI.IsChecked = DirectCast(oDataRow("bALI"), System.Boolean)
            End If
        End If
        GetStateTax()
        sRowTag = oTask.Row.Tag.Split("|"c)
        txtDescription.Text = mp_oParent.GetDescription(CType(sRowTag(2), System.Int32))
        picCarType.Source = GetImageSource(g_GetAppLocation() & "\CarRental\Small\" & txtDescription.Text & ".jpg")
        txtRate.Text = g_Format(CType(sRowTag(1), System.Double), "0.00")
        txtACRISS1.Text = GetACRISSDescription(1, sRowTag(0).Substring(0, 1))
        txtACRISS2.Text = GetACRISSDescription(2, sRowTag(0).Substring(1, 1))
        txtACRISS3.Text = GetACRISSDescription(3, sRowTag(0).Substring(2, 1))
        txtACRISS4.Text = GetACRISSDescription(4, sRowTag(0).Substring(3, 1))
        CalculateRate()
    End Sub

    Private Sub GetStateTax()
        Dim oTask As clsTask = GetCurrentTask()
        Dim sState As String = ""
        Dim dTax As Double = 0
        dTax = mp_oParent.GetStateTax(oTask, sState)
        lblTax.Content = sState & " State Tax (" & g_Format(dTax * 100, "0") & "%):"
        lblTax.Tag = dTax
    End Sub

    Private Sub CalculateRate()
        Dim fFactor As Single = 0
        Dim sRowTag As String()
        Dim lRate As Single = 0
        Dim fSubTotal As Single = 0
        Dim fOptions As Single = 0
        Dim oTask As clsTask = GetCurrentTask()
        fFactor = CType(mp_oParent.ActiveGanttVBWCtl1.MathLib.DateTimeDiff(E_INTERVAL.IL_HOUR, GetStartDate(), GetEndDate()) / 24, System.Single)
        If chkGPS.IsChecked = True Then
            txtGPS.Text = g_Format(DirectCast(chkGPS.Tag, System.Double) * fFactor, "0.00")
            txtGPS.Tag = g_Format(DirectCast(chkGPS.Tag, System.Double) * fFactor, "0.00")
        Else
            txtGPS.Text = ""
            txtGPS.Tag = 0
        End If
        If chkLDW.IsChecked = True Then
            txtLDW.Text = g_Format(DirectCast(chkLDW.Tag, System.Double) * fFactor, "0.00")
            txtLDW.Tag = g_Format(DirectCast(chkLDW.Tag, System.Double) * fFactor, "0.00")
        Else
            txtLDW.Text = ""
            txtLDW.Tag = 0
        End If
        If chkPAI.IsChecked = True Then
            txtPAI.Text = g_Format(DirectCast(chkPAI.Tag, System.Double) * fFactor, "0.00")
            txtPAI.Tag = g_Format(DirectCast(chkPAI.Tag, System.Double) * fFactor, "0.00")
        Else
            txtPAI.Text = ""
            txtPAI.Tag = 0
        End If
        If chkPEP.IsChecked = True Then
            txtPEP.Text = g_Format(DirectCast(chkPEP.Tag, System.Double) * fFactor, "0.00")
            txtPEP.Tag = g_Format(DirectCast(chkPEP.Tag, System.Double) * fFactor, "0.00")
        Else
            txtPEP.Text = ""
            txtPEP.Tag = 0
        End If
        If chkALI.IsChecked = True Then
            txtALI.Text = g_Format(DirectCast(chkALI.Tag, System.Double) * fFactor, "0.00")
            txtALI.Tag = g_Format(DirectCast(chkALI.Tag, System.Double) * fFactor, "0.00")
        Else
            txtALI.Text = ""
            txtALI.Tag = 0
        End If
        sRowTag = oTask.Row.Tag.Split("|"c)
        lRate = CType(sRowTag(1), System.Single)
        txtERF.Text = g_Format(CType(txtERF.Tag, System.Double) * fFactor, "0.00")
        txtWTB.Text = g_Format(CType(txtWTB.Tag, System.Double) * fFactor, "0.00")
        txtRCFC.Text = g_Format(CType(txtRCFC.Tag, System.Double) * fFactor, "0.00")
        txtVLF.Text = g_Format(CType(txtVLF.Tag, System.Double) * fFactor, "0.00")
        txtCRF.Text = g_Format(CType(txtCRF.Tag, System.Double) * lRate * fFactor, "0.00")
        txtSurcharge.Tag = (CType(txtERF.Tag, System.Double) * fFactor) + (CType(txtWTB.Tag, System.Double) * fFactor) + (CType(txtRCFC.Tag, System.Double) * fFactor) + (CType(txtVLF.Tag, System.Double) * fFactor) + (CType(txtCRF.Tag, System.Double) * lRate * fFactor)
        txtSurcharge.Text = g_Format(CType(txtSurcharge.Tag, System.Double), "0.00")

        fOptions = CType(txtGPS.Tag, System.Single) + CType(txtLDW.Tag, System.Single) + CType(txtPAI.Tag, System.Single) + CType(txtPEP.Tag, System.Single) + CType(txtALI.Tag, System.Single)
        fSubTotal = CType(txtSurcharge.Tag, System.Single) + (lRate * fFactor)

        txtTax.Tag = fSubTotal * DirectCast(lblTax.Tag, System.Double)
        txtTax.Text = g_Format(DirectCast(txtTax.Tag, System.Double), "0.00")

        txtEstimatedTotal.Tag = fSubTotal + CType(txtTax.Tag, System.Single) + fOptions
        txtEstimatedTotal.Text = g_Format(CType(txtEstimatedTotal.Tag, System.Double), "0.00")
    End Sub

    Public Function GetACRISSDescription(ByVal sPosition As Integer, ByVal sLetter As String) As String
        Dim sReturn As String = "GetACRISSDescription Error"
        Dim oDataRow As DataRow = Nothing
        If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
                Dim oReader As OleDbDataReader = Nothing
                oReader = g_DST_ACCESS_ReturnReader("SELECT * FROM tb_CR_ACRISS_Codes WHERE [Letter] = '" & sLetter & "' AND [Position] = " & sPosition, oConn)
                If oReader.Read = True Then
                    sReturn = DirectCast(oReader.Item("Description"), System.String)
                End If
                oReader.Close()
            End Using
            Return sReturn
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            oDataRow = g_DST_XML_FindRow(mp_oParent.mp_otb_CR_ACRISS_Codes, "Letter", sLetter)
            sReturn = DirectCast(oDataRow("Description"), System.String)
            Return sReturn
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            oDataRow = g_DST_NONE_FindRow(mp_oParent.mp_otb_CR_ACRISS_Codes, "Letter", sLetter)
            sReturn = DirectCast(oDataRow("Description"), System.String)
            Return sReturn
        End If
        Return sReturn
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

    Private Sub chkGPS_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles chkGPS.Click
        CalculateRate()
    End Sub

    Private Sub chkLDW_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles chkLDW.Click
        CalculateRate()
    End Sub

    Private Sub chkPAI_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles chkPAI.Click
        CalculateRate()
    End Sub

    Private Sub chkPEP_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles chkPEP.Click
        CalculateRate()
    End Sub

    Private Sub chkALI_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles chkALI.Click
        CalculateRate()
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdOK.Click
        Dim oDataRow As DataRow = Nothing
        Dim oTask As clsTask = GetCurrentTask()
        If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            Dim oDB As clsDB = Nothing
            oDB = New clsDB()
            oDB.AddParameter("lRowID", oTask.RowKey.Replace("K", ""), clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("yMode", CType(mp_oParent.Mode, System.Int32), clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("sCity", txtCity.Text, clsDB.ParamType.PT_STRING)
            oDB.AddParameter("sName", txtName.Text, clsDB.ParamType.PT_STRING)
            oDB.AddParameter("sState", txtState.Text, clsDB.ParamType.PT_STRING)
            oDB.AddParameter("sPhone", txtPhone.Text, clsDB.ParamType.PT_STRING)
            oDB.AddParameter("sMobile", txtMobile.Text, clsDB.ParamType.PT_STRING)
            oDB.AddParameter("sAddress", txtAddress.Text, clsDB.ParamType.PT_STRING)
            oDB.AddParameter("sZIP", txtZIP.Text, clsDB.ParamType.PT_STRING)
            oDB.AddParameter("dtPickUp", GetStartDate(), clsDB.ParamType.PT_DATE)
            oDB.AddParameter("dtReturn", GetEndDate(), clsDB.ParamType.PT_DATE)
            oDB.AddParameter("cRate", txtRate.Text, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("cGPS", chkGPS.Tag, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("cLDW", chkLDW.Tag, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("cPAI", chkPAI.Tag, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("cPEP", chkPEP.Tag, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("cALI", chkALI.Tag, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("cERF", txtERF.Tag, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("dCRF", txtCRF.Tag, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("cRCFC", txtRCFC.Tag, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("cWTB", txtWTB.Tag, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("cVLF", txtVLF.Tag, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("dTax", lblTax.Tag, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("cEstimatedTotal", txtEstimatedTotal.Tag, clsDB.ParamType.PT_NUMERIC)
            oDB.AddParameter("bGPS", chkGPS.IsChecked, clsDB.ParamType.PT_BOOL)
            oDB.AddParameter("bFSO", chkFSO.IsChecked, clsDB.ParamType.PT_BOOL)
            oDB.AddParameter("bLDW", chkLDW.IsChecked, clsDB.ParamType.PT_BOOL)
            oDB.AddParameter("bPAI", chkPAI.IsChecked, clsDB.ParamType.PT_BOOL)
            oDB.AddParameter("bPEP", chkPEP.IsChecked, clsDB.ParamType.PT_BOOL)
            oDB.AddParameter("bALI", chkALI.IsChecked, clsDB.ParamType.PT_BOOL)
            If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
                oTask.Key = "K" & oDB.ExecuteInsert("tb_CR_Rentals")
            ElseIf mp_yDialogMode = PRG_DIALOGMODE.DM_EDIT Then
                oDB.ExecuteUpdate("tb_CR_Rentals", "lTaskID = " & oTask.Key.Replace("K", ""))
            End If
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
                Dim lTaskID As Integer = g_DST_XML_AutoIncrementValue(mp_oParent.mp_otb_CR_Rentals, "lTaskID")
                oDataRow = mp_oParent.mp_otb_CR_Rentals.Tables(1).NewRow()
                oDataRow("lTaskID") = lTaskID
                oTask.Key = "K" & lTaskID.ToString()
            ElseIf mp_yDialogMode = PRG_DIALOGMODE.DM_EDIT Then
                oDataRow = mp_oParent.mp_otb_CR_Rentals.Tables(1).Rows.Find(oTask.Key.Replace("K", ""))
            End If
            oDataRow("lRowID") = oTask.RowKey.Replace("K", "")
            oDataRow("yMode") = CType(CType(mp_oParent.Mode, System.Int32), System.String)
            oDataRow("sCity") = txtCity.Text
            oDataRow("sName") = txtName.Text
            oDataRow("sState") = txtState.Text
            oDataRow("sPhone") = txtPhone.Text
            oDataRow("sMobile") = txtMobile.Text
            oDataRow("sAddress") = txtAddress.Text
            oDataRow("sZIP") = txtZIP.Text
            oDataRow("dtPickUp") = GetStartDate().DateTimePart
            oDataRow("dtReturn") = GetEndDate().DateTimePart
            oDataRow("cRate") = txtRate.Text
            oDataRow("cGPS") = chkGPS.Tag
            oDataRow("cLDW") = chkLDW.Tag
            oDataRow("cPAI") = chkPAI.Tag
            oDataRow("cPEP") = chkPEP.Tag
            oDataRow("cALI") = chkALI.Tag
            oDataRow("cERF") = txtERF.Tag
            oDataRow("dCRF") = txtCRF.Tag
            oDataRow("cRCFC") = txtRCFC.Tag
            oDataRow("cWTB") = txtWTB.Tag
            oDataRow("cVLF") = txtVLF.Tag
            oDataRow("dTax") = lblTax.Tag
            oDataRow("cEstimatedTotal") = txtEstimatedTotal.Tag
            oDataRow("bGPS") = chkGPS.IsChecked
            oDataRow("bFSO") = chkFSO.IsChecked
            oDataRow("bLDW") = chkLDW.IsChecked
            oDataRow("bPAI") = chkPAI.IsChecked
            oDataRow("bPEP") = chkPEP.IsChecked
            oDataRow("bALI") = chkALI.IsChecked
            If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
                mp_oParent.mp_otb_CR_Rentals.Tables(1).Rows.Add(oDataRow)
            End If
            mp_oParent.mp_otb_CR_Rentals.WriteXml(g_GetAppLocation() & "\CR_XML\tb_CR_Rentals.xml")
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
                Dim lTaskID As Integer = g_DST_NONE_AutoIncrementValue(mp_oParent.mp_otb_CR_Rentals, "lTaskID")
                oDataRow = mp_oParent.mp_otb_CR_Rentals.Tables(0).NewRow()
                oDataRow("lTaskID") = lTaskID
                oTask.Key = "K" & lTaskID.ToString()
            ElseIf mp_yDialogMode = PRG_DIALOGMODE.DM_EDIT Then
                oDataRow = mp_oParent.mp_otb_CR_Rentals.Tables(0).Rows.Find(oTask.Key.Replace("K", ""))
            End If
            oDataRow("lRowID") = oTask.RowKey.Replace("K", "")
            oDataRow("yMode") = CType(CType(mp_oParent.Mode, System.Int32), System.String)
            oDataRow("sCity") = txtCity.Text
            oDataRow("sName") = txtName.Text
            oDataRow("sState") = txtState.Text
            oDataRow("sPhone") = txtPhone.Text
            oDataRow("sMobile") = txtMobile.Text
            oDataRow("sAddress") = txtAddress.Text
            oDataRow("sZIP") = txtZIP.Text
            oDataRow("dtPickUp") = GetStartDate().DateTimePart
            oDataRow("dtReturn") = GetEndDate().DateTimePart
            oDataRow("cRate") = txtRate.Text
            oDataRow("cGPS") = chkGPS.Tag
            oDataRow("cLDW") = chkLDW.Tag
            oDataRow("cPAI") = chkPAI.Tag
            oDataRow("cPEP") = chkPEP.Tag
            oDataRow("cALI") = chkALI.Tag
            oDataRow("cERF") = txtERF.Tag
            oDataRow("dCRF") = txtCRF.Tag
            oDataRow("cRCFC") = txtRCFC.Tag
            oDataRow("cWTB") = txtWTB.Tag
            oDataRow("cVLF") = txtVLF.Tag
            oDataRow("dTax") = lblTax.Tag
            oDataRow("cEstimatedTotal") = txtEstimatedTotal.Tag
            oDataRow("bGPS") = chkGPS.IsChecked
            oDataRow("bFSO") = chkFSO.IsChecked
            oDataRow("bLDW") = chkLDW.IsChecked
            oDataRow("bPAI") = chkPAI.IsChecked
            oDataRow("bPEP") = chkPEP.IsChecked
            oDataRow("bALI") = chkALI.IsChecked
            If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
                mp_oParent.mp_otb_CR_Rentals.Tables(0).Rows.Add(oDataRow)
            End If
        End If

        oTask.Text = txtName.Text & vbCrLf & "Phone: " & txtPhone.Text & vbCrLf & "Estimated Total: " & txtEstimatedTotal.Text & " USD"
        oTask.Tag = CType(CType(mp_oParent.Mode, System.Int32), System.String)
        mp_oParent.ActiveGanttVBWCtl1.Redraw()
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdCancel.Click
        If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
            mp_oParent.ActiveGanttVBWCtl1.Tasks.Remove(mp_oParent.ActiveGanttVBWCtl1.Tasks.Count.ToString)
        End If
        Me.Close()
    End Sub

    Private Sub SetStartDate(ByVal dtDate As AGVBW.DateTime)
        txtSDYear.Text = dtDate.Year
        txtSDMonth.Text = dtDate.Month
        txtSDDay.Text = dtDate.Day
        txtSDHours.Text = dtDate.Hour
        txtSDMinutes.Text = dtDate.Minute
        txtSDSeconds.Text = dtDate.Second
    End Sub

    Private Sub SetEndDate(ByVal dtDate As AGVBW.DateTime)
        txtEDYear.Text = dtDate.Year
        txtEDMonth.Text = dtDate.Month
        txtEDDay.Text = dtDate.Day
        txtEDHours.Text = dtDate.Hour
        txtEDMinutes.Text = dtDate.Minute
        txtEDSeconds.Text = dtDate.Second
    End Sub

    Private Function GetStartDate() As AGVBW.DateTime
        Dim dtDate As New AGVBW.DateTime(txtSDYear.Text, txtSDMonth.Text, txtSDDay.Text, txtSDHours.Text, txtSDMinutes.Text, txtSDSeconds.Text)
        Return dtDate
    End Function

    Private Function GetEndDate() As AGVBW.DateTime
        Dim dtDate As New AGVBW.DateTime(txtEDYear.Text, txtEDMonth.Text, txtEDDay.Text, txtEDHours.Text, txtEDMinutes.Text, txtEDSeconds.Text)
        Return dtDate
    End Function

    Private Function GetCurrentTask() As clsTask
        Dim oReturnTask As clsTask = Nothing
        If mp_yDialogMode = PRG_DIALOGMODE.DM_EDIT Then
            oReturnTask = mp_oParent.ActiveGanttVBWCtl1.Tasks.Item("K" & mp_sTaskID)
        ElseIf mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
            oReturnTask = mp_oParent.ActiveGanttVBWCtl1.Tasks.Item(mp_oParent.ActiveGanttVBWCtl1.Tasks.Count.ToString())
        End If
        Return oReturnTask
    End Function

End Class
