Option Explicit On

Imports System.Data.OleDb
Imports System.Data


Module DST_ACCESS

    Public Function g_DST_ACCESS_GetDatabaseLocation() As String
        Return AppDomain.CurrentDomain.BaseDirectory.Replace("\bin\", "") & "\ActiveGanttExamples.mdb"
    End Function

    Public Function g_DST_ACCESS_GetConnectionString() As String
        Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_DST_ACCESS_GetDatabaseLocation()
    End Function

    Public Sub g_DST_ACCESS_FillComboBox(ByRef oComboBox As ComboBox, ByVal sSQL As String, ByVal sValueMember As String, ByVal sDisplayMember As String)
        Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
            Dim oAdapter As New OleDbDataAdapter()
            Dim oDataTable As New DataTable()
            oAdapter.SelectCommand = New OleDbCommand(sSQL, oConn)
            oAdapter.Fill(oDataTable)
            oComboBox.ItemsSource = oDataTable.DefaultView
            oComboBox.DisplayMemberPath = sDisplayMember
            oComboBox.SelectedValuePath = sValueMember
            oConn.Close()
        End Using
    End Sub

    Public Function g_DST_ACCESS_ConvertDate(ByVal dtDate As AGVBW.DateTime) As String
        Dim sReturn As String = ""
        sReturn = "#" & dtDate.ToString("yyyy-MM-dd HH:mm:ss") & "#"
        Return sReturn
    End Function

    Public Function g_DST_ACCESS_InsertWithID(ByVal sSQL As String) As String
        Dim sReturn As String = ""
        Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
            Dim oComm As OleDbCommand = Nothing
            Dim oReader As OleDbDataReader = Nothing
            oConn.Open()
            oComm = New OleDbCommand(sSQL, oConn)
            oComm.CommandText = sSQL
            oComm.ExecuteNonQuery()
            oComm = New OleDbCommand("SELECT @@IDENTITY AS NewID", oConn)
            oReader = oComm.ExecuteReader()
            If oReader.Read = True Then
                sReturn = System.Convert.ToString(oReader.Item("NewID"))
            End If
            oReader.Close()
            oConn.Close()
        End Using

        Return sReturn
    End Function

    Public Sub g_DST_ACCESS_ExecuteNonQuery(ByVal sSQL As String)
        Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
            Dim oComm As OleDbCommand = Nothing
            oConn.Open()
            oComm = New OleDbCommand(sSQL, oConn)
            oComm.ExecuteNonQuery()
            oConn.Close()
        End Using
    End Sub

    Public Function g_DST_ACCESS_ReturnReader(ByVal sSQL As String, ByRef oConn As OleDbConnection) As OleDbDataReader
        Dim oComm As OleDbCommand = Nothing
        Dim oReader As OleDbDataReader = Nothing
        If oConn.State = ConnectionState.Closed Then
            oConn.ConnectionString = g_DST_ACCESS_GetConnectionString()
            oConn.Open()
        End If
        oComm = New OleDbCommand(sSQL, oConn)
        oReader = oComm.ExecuteReader()
        Return oReader
    End Function

End Module
