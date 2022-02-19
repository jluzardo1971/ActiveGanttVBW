Option Explicit On

Imports System.Data

Module DST_XML

    Public Function g_DST_XML_AutoIncrementValue(ByRef oDataSet As DataSet, ByVal sColumnName As String) As Integer
        Dim oDataRow As DataRow = Nothing
        Dim lMax As Integer = 0
        For Each oDataRow In oDataSet.Tables(1).Rows()
            If DirectCast(oDataRow(sColumnName), System.Int32) > lMax Then
                lMax = DirectCast(oDataRow(sColumnName), System.Int32)
            End If
        Next
        Return lMax + 1
    End Function

    Public Function g_DST_XML_FindRow(ByRef oDataSet As DataSet, ByVal sColumnName As String, ByVal sColumnValue As String) As DataRow
        Dim oDataRow As DataRow = Nothing
        For Each oDataRow In oDataSet.Tables(1).Rows()
            If DirectCast(oDataRow(sColumnName), System.String) = sColumnValue Then
                Return oDataRow
            End If
        Next
        Return Nothing
    End Function

    Public Sub g_DST_XML_FillComboBox(ByRef oComboBox As ComboBox, ByRef oDataTable As DataTable, ByVal sValueMember As String, ByVal sDisplayMember As String, Optional ByVal sFilter As String = "")
        If sFilter.Length() > 0 Then
            oDataTable.DefaultView.RowFilter = sFilter
        End If
        oComboBox.ItemsSource = oDataTable.DefaultView
        oComboBox.DisplayMemberPath = sDisplayMember
        oComboBox.SelectedValuePath = sValueMember
    End Sub

End Module
