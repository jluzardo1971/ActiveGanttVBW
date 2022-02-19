Imports System.Data.OleDb

Public Class clsDB

    Private mp_asFieldNames() As String
    Private mp_asParams() As String

    Private mp_oOLEDBConnection As OleDbConnection = Nothing
    Private mp_oOLEDBDataReader As OleDbDataReader
    Private mp_bReaderActive As Boolean = False

    Public Enum ParamType
        PT_STRING = 0
        PT_STRING_EMPTYISNULL = 1
        PT_NUMERIC = 2
        PT_NUMERIC_ZEROISNULL = 3
        PT_BOOL = 4
        PT_DATE = 5
        PT_NULLVALUE = 6
        PT_UNIQUEIDENTIFIER = 7
        PT_DATE_EMPTYISNULL = 8
    End Enum

    Public Sub New()
        ReDim mp_asFieldNames(0)
        ReDim mp_asParams(0)
        mp_oOLEDBConnection = New OleDbConnection()
        mp_oOLEDBConnection.ConnectionString = DST_ACCESS.g_DST_ACCESS_GetConnectionString()
        mp_oOLEDBConnection.Open()
    End Sub

    Public Sub InitReader(ByVal sSQL As String)
        Dim oComm As OleDbCommand
        oComm = New OleDbCommand(sSQL, mp_oOLEDBConnection)
        mp_oOLEDBDataReader = oComm.ExecuteReader()
        mp_bReaderActive = mp_oOLEDBDataReader.Read()
    End Sub

    Private Sub InitReaderFill(ByVal sSQL As String)
        Dim oComm As OleDbCommand
        oComm = New OleDbCommand(sSQL, mp_oOLEDBConnection)
        mp_oOLEDBDataReader = oComm.ExecuteReader()
        mp_bReaderActive = False
    End Sub

    Public Sub CloseReader()
        mp_oOLEDBDataReader.Close()
    End Sub

    Public Function IsDBNull(ByVal sFieldName As String) As Boolean
        Dim bReturn As Boolean = False
        If mp_bReaderActive = False Then Return False
        bReturn = DBNull.Value.Equals(mp_oOLEDBDataReader.Item(sFieldName))
        Return bReturn
    End Function


    Public Sub Read(ByVal oControl As CheckBox, ByVal sFieldName As String)
        If mp_bReaderActive = False Then Return
        If DBNull.Value.Equals(mp_oOLEDBDataReader.Item(sFieldName)) = True Then
            oControl.IsChecked = False
        Else
            oControl.IsChecked = CBool(mp_oOLEDBDataReader.Item(sFieldName))
        End If
    End Sub

    Public Sub Write(ByVal oControl As CheckBox, ByVal sFieldName As String)
        AddParameter(sFieldName, oControl.IsChecked, ParamType.PT_BOOL)
    End Sub

    Public Sub Write(ByVal oControl As TextBox, ByVal sFieldName As String, Optional ByVal v_sType As ParamType = ParamType.PT_STRING_EMPTYISNULL)
        AddParameter(sFieldName, oControl.Text, v_sType)
    End Sub

    Public Sub Write(ByVal oControl As Label, ByVal sFieldName As String, Optional ByVal v_sType As ParamType = ParamType.PT_STRING_EMPTYISNULL)
        AddParameter(sFieldName, oControl.Content, v_sType)
    End Sub

    Public Sub Write(ByVal oControl As ComboBox, ByVal sFieldName As String, Optional ByVal v_sType As ParamType = ParamType.PT_NUMERIC)
        AddParameter(sFieldName, oControl.SelectedValue, v_sType)
    End Sub

    Public Sub Read(ByVal oControl As TextBox, ByVal sFieldName As String)
        If mp_bReaderActive = False Then Return
        If DBNull.Value.Equals(mp_oOLEDBDataReader.Item(sFieldName)) = True Then
            oControl.Text = ""
        Else
            oControl.Text = mp_oOLEDBDataReader.Item(sFieldName)
        End If
    End Sub

    Public Sub Read(ByRef oControl As String, ByVal sFieldName As String)
        If mp_bReaderActive = False Then Return
        If DBNull.Value.Equals(mp_oOLEDBDataReader.Item(sFieldName)) = True Then
            oControl = ""
        Else
            oControl = mp_oOLEDBDataReader.Item(sFieldName)
        End If
    End Sub

    Public Function Read(ByVal sFieldName As String) As String
        Dim sReturn As String = ""
        If mp_bReaderActive = False Then Return sReturn
        If DBNull.Value.Equals(mp_oOLEDBDataReader.Item(sFieldName)) = True Then
            sReturn = ""
        Else
            sReturn = mp_oOLEDBDataReader.Item(sFieldName)
        End If
        Return sReturn
    End Function

    Public Sub Read(ByVal oControl As Label, ByVal sFieldName As String)
        If mp_bReaderActive = False Then Return
        If DBNull.Value.Equals(mp_oOLEDBDataReader.Item(sFieldName)) = True Then
            oControl.Content = ""
        Else
            oControl.Content = mp_oOLEDBDataReader.Item(sFieldName)
        End If
    End Sub

    Public Sub Read(ByVal oControl As ComboBox, ByVal sFieldName As String)
        If mp_bReaderActive = False Then Return
        If DBNull.Value.Equals(mp_oOLEDBDataReader.Item(sFieldName)) = True Then
            oControl.SelectedValue = ""
        Else
            oControl.SelectedValue = mp_oOLEDBDataReader.Item(sFieldName)
        End If
    End Sub

    Public Sub Fill(ByVal sSQL As String, ByVal oControl As ComboBox)

    End Sub

    Public Sub Fill(ByVal sSQL As String, ByVal oControl As ComboBox, ByVal sDataValueField As String, ByVal sDataTextField As String)

    End Sub


    Public Sub AddParameter(ByVal v_sFieldName As String, ByVal oParam As Object, ByVal v_sType As ParamType)
        '// Inserts parameters for future inserts or updates
        ReDim Preserve mp_asFieldNames(UBound(mp_asFieldNames, 1) + 1)
        ReDim Preserve mp_asParams(UBound(mp_asParams, 1) + 1)
        mp_asFieldNames(UBound(mp_asFieldNames, 1)) = v_sFieldName
        Dim sParam As String = ""
        Dim dtParam As AGVBW.DateTime = New AGVBW.DateTime()
        Select Case v_sType
            Case ParamType.PT_STRING
                If Trim$(oParam) = "" Then
                    mp_asParams(UBound(mp_asParams, 1)) = "''"
                Else
                    oParam = Replace(oParam, "'", "''")
                    mp_asParams(UBound(mp_asParams, 1)) = "'" & oParam & "'"
                End If
            Case ParamType.PT_STRING_EMPTYISNULL
                If Trim$(oParam) = "" Then
                    mp_asParams(UBound(mp_asParams, 1)) = "NULL"
                Else
                    oParam = Replace(oParam, "'", "''")
                    mp_asParams(UBound(mp_asParams, 1)) = "'" & oParam & "'"
                End If
            Case ParamType.PT_UNIQUEIDENTIFIER
                If Trim$(oParam) = "" Then
                    mp_asParams(UBound(mp_asParams, 1)) = "NULL"
                Else
                    mp_asParams(UBound(mp_asParams, 1)) = "'" & oParam & "'"
                End If
            Case ParamType.PT_NUMERIC
                If IsNumeric(oParam) Then
                    mp_asParams(UBound(mp_asParams, 1)) = oParam
                Else
                    mp_asParams(UBound(mp_asParams, 1)) = "NULL"
                End If
            Case ParamType.PT_NUMERIC_ZEROISNULL
                If oParam = 0 Then
                    mp_asParams(UBound(mp_asParams, 1)) = "NULL"
                Else
                    mp_asParams(UBound(mp_asParams, 1)) = oParam
                End If
            Case ParamType.PT_BOOL
                If oParam Then
                    mp_asParams(UBound(mp_asParams, 1)) = "-1"
                Else
                    mp_asParams(UBound(mp_asParams, 1)) = "0"
                End If
            Case ParamType.PT_DATE_EMPTYISNULL
                sParam = DirectCast(oParam, System.String)
                If sParam.Length = 0 Then
                    mp_asParams(UBound(mp_asParams, 1)) = "NULL"
                Else
                    dtParam = oParam
                    mp_asParams(UBound(mp_asParams, 1)) = "#" & dtParam.ToString("yyyy-MM-dd HH:mm:ss") & "#"
                End If
            Case ParamType.PT_DATE
                dtParam = oParam
                mp_asParams(UBound(mp_asParams, 1)) = "#" & dtParam.ToString("yyyy-MM-dd HH:mm:ss") & "#"
            Case ParamType.PT_NULLVALUE
                mp_asParams(UBound(mp_asParams, 1)) = "NULL"
        End Select
    End Sub

    Public Function Insert(ByVal v_sTableName As String) As String
        '// SQL syntax independent insert Statement
        Dim iIndex As Integer
        Dim sSQL As String
        sSQL = "INSERT INTO " & v_sTableName & " ("
        For iIndex = 1 To UBound(mp_asFieldNames)
            sSQL = sSQL & "[" & mp_asFieldNames(iIndex) & "],"
        Next iIndex
        sSQL = Left$(sSQL, Len(sSQL) - 1)
        sSQL = sSQL & ") VALUES ("
        For iIndex = 1 To UBound(mp_asParams)
            sSQL = sSQL & mp_asParams(iIndex) & ","
        Next iIndex
        sSQL = Left$(sSQL, Len(sSQL) - 1)
        sSQL = sSQL & ")"
        ReDim mp_asFieldNames(0)
        ReDim mp_asParams(0)
        Return sSQL
    End Function

    Public Function Update(ByVal v_sTableName As String, ByVal v_sWHERE As String) As String
        '// SQL syntax independent update Statement
        Dim iIndex As Integer
        Dim sSQL As String
        sSQL = "UPDATE " & v_sTableName & " SET "
        For iIndex = 1 To UBound(mp_asFieldNames)
            sSQL = sSQL & "[" & mp_asFieldNames(iIndex) & "]=" & mp_asParams(iIndex) & ","
        Next iIndex
        sSQL = Left$(sSQL, Len(sSQL) - 1)
        sSQL = sSQL & " WHERE " & v_sWHERE
        ReDim mp_asFieldNames(0)
        ReDim mp_asParams(0)
        Return sSQL
    End Function

    Public Function ExecuteInsert(ByVal sTableName As String) As String
        Return DST_ACCESS.g_DST_ACCESS_InsertWithID(Insert(sTableName))
        Return ""
    End Function

    Public Sub ExecuteUpdate(ByVal sTableName As String, ByVal sWHERE_WITHOUT_WHERE As String)
        DST_ACCESS.g_DST_ACCESS_ExecuteNonQuery(Update(sTableName, sWHERE_WITHOUT_WHERE))
    End Sub

    Public Function Delete(ByVal v_sTableName As String, Optional ByVal v_sWHERE As String = "") As String
        '// SQL syntax independent Delete Statement
        Dim sSQL As String
        sSQL = "DELETE FROM " & v_sTableName
        If v_sWHERE <> "" Then
            sSQL = sSQL & " WHERE " & v_sWHERE
        End If
        Return sSQL
    End Function

End Class
