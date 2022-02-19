Option Explicit On 
Imports System.Xml

Partial Friend Class clsXML

    Private xDoc As System.Xml.XmlDocument
    Private oControlElement As System.Xml.XmlElement
    Private oFontElement As System.Xml.XmlElement
    Private oDateTimeElement As System.Xml.XmlElement
    Private mp_sObject As String
    Private mp_yLevel As PE_LEVEL
    Private mp_bSupportOptional As Boolean = False
    Private mp_bBoolsAreNumeric As Boolean = False

    Private Enum PE_LEVEL
        LVL_CONTROL = 0
        LVL_FONT = 5
        LVL_DATETIME = 6
    End Enum

    Friend Property SupportOptional() As Boolean
        Get
            Return mp_bSupportOptional
        End Get
        Set(ByVal value As Boolean)
            mp_bSupportOptional = value
        End Set
    End Property

    Friend Property BoolsAreNumeric() As Boolean
        Get
            Return mp_bBoolsAreNumeric
        End Get
        Set(ByVal value As Boolean)
            mp_bBoolsAreNumeric = value
        End Set
    End Property

    Friend Sub New(ByVal sObject As String)
        xDoc = New System.Xml.XmlDocument()
        mp_sObject = sObject
    End Sub

    Friend Sub InitializeWriter()
        xDoc.LoadXml("<" & mp_sObject & "></" & mp_sObject & ">")
        oControlElement = GetDocumentElement(mp_sObject, 0)
        mp_yLevel = PE_LEVEL.LVL_CONTROL
    End Sub

    Friend Sub InitializeReader()
        oControlElement = GetDocumentElement(mp_sObject, 0)
        mp_yLevel = PE_LEVEL.LVL_CONTROL
    End Sub

    Friend ReadOnly Property GetDocument() As System.Xml.XmlDocument
        Get
            Return xDoc
        End Get
    End Property

    Friend Sub AddAttribute(ByVal sName As String, ByVal sValue As String)
        Dim oAttribute As Xml.XmlAttribute = xDoc.CreateAttribute(sName)
        oAttribute.Value = sValue
        xDoc.DocumentElement.Attributes.Append(oAttribute)
    End Sub

    Friend Sub WriteXML(ByVal sPath As String)
        Dim oWriter As XmlTextWriter = New XmlTextWriter(sPath, System.Text.Encoding.UTF8)
        oWriter.IndentChar = ControlChars.Tab
        oWriter.Formatting = Formatting.Indented
        oWriter.WriteStartDocument()

        xDoc.Save(oWriter)
        oWriter.WriteEndDocument()
        oWriter.Close()
    End Sub

    Friend Sub ReadXML(ByVal sPath As String)
        xDoc.Load(sPath)
    End Sub

    Private Function ParentElement() As System.Xml.XmlElement
        Select Case mp_yLevel
            Case PE_LEVEL.LVL_CONTROL
                Return oControlElement
            Case PE_LEVEL.LVL_FONT
                Return oFontElement
            Case PE_LEVEL.LVL_DATETIME
                Return oDateTimeElement
        End Select
        Return Nothing
    End Function

    Private Function mp_oCreateEmptyDOMElement(ByVal sElementName As String) As System.Xml.XmlElement
        Dim oNodeBuff As System.Xml.XmlElement
        oNodeBuff = xDoc.CreateElement(sElementName)
        ParentElement.AppendChild(oNodeBuff)
        Return oNodeBuff
    End Function

    Private Function GetDocumentElement(ByVal TagName As String, ByVal lIndex As Integer) As System.Xml.XmlElement
        Return xDoc.GetElementsByTagName(TagName).Item(lIndex)
    End Function

    Friend Function GetXML() As String
        Return xDoc.InnerXml
    End Function

    Friend Sub SetXML(ByVal sXML As String)
        If mp_bSupportOptional = False Then
            xDoc.LoadXml(sXML)
        Else
            If sXML.Length > 0 Then
                xDoc.LoadXml(sXML)
            End If
        End If
    End Sub

    Friend Function ReadObject(ByVal sObjectName As String) As String
        If mp_bSupportOptional = False Then
            Return ParentElement().GetElementsByTagName(sObjectName).Item(0).OuterXml()
        Else
            If ParentElement() Is Nothing Then
                Return ""
            End If
            If ParentElement().GetElementsByTagName(sObjectName).Count > 0 Then
                Return ParentElement().GetElementsByTagName(sObjectName).Item(0).OuterXml()
            Else
                Return ""
            End If
        End If
    End Function

    Friend Function GetFirstCollectionItem(ByVal sItemName As String) As Integer
        Dim lIndex As Integer
        For lIndex = 1 To ReadCollectionCount()
            If ParentElement().ChildNodes.Item(lIndex - 1).Name = sItemName Then
                Return lIndex
            End If
        Next
        Return ReadCollectionCount() + 1
    End Function

    Friend Function GetLastCollectionItem(ByVal sItemName As String) As Integer
        Dim lIndex As Integer
        For lIndex = ReadCollectionCount() To 1 Step -1
            If ParentElement().ChildNodes.Item(lIndex - 1).Name = sItemName Then
                Return lIndex
            End If
        Next
        Return 0
    End Function

    Friend Function ReadCollectionObject(ByVal lIndex As Integer) As String
        If mp_bSupportOptional = False Then
            Return ParentElement().ChildNodes.Item(lIndex - 1).OuterXml()
        Else
            If ParentElement() Is Nothing Or lIndex = 0 Then
                Return ""
            End If
            If ParentElement.ChildNodes.Count > 0 Then
                Return ParentElement().ChildNodes.Item(lIndex - 1).OuterXml()
            Else
                Return ""
            End If
        End If
    End Function

    Friend Function GetCollectionObjectName(ByVal lIndex As Integer) As String
        Return ParentElement().ChildNodes.Item(lIndex - 1).Name
    End Function

    Friend Function ReadCollectionCount() As Integer
        If mp_bSupportOptional = False Then
            Return ParentElement().ChildNodes.Count
        Else
            If ParentElement() Is Nothing Then
                Return 0
            Else
                Return ParentElement().ChildNodes.Count
            End If
        End If
    End Function

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As Integer)
        sElementValue = lReadProperty(sElementName, sElementValue)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef iElementValue As Short)
        iElementValue = iReadProperty(sElementName, iElementValue)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As String)
        sElementValue = sReadProperty(sElementName, sElementValue)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef bElementValue As Boolean)
        bElementValue = bReadProperty(sElementName, bElementValue)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef dtElementValue As System.DateTime)
        dtElementValue = dtReadProperty(sElementName, dtElementValue)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef fElementValue As Single)
        fElementValue = fReadProperty(sElementName, fElementValue)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As Byte)
        Dim yBuff As Short = 0
        yBuff = yReadProperty(sElementName, sElementValue)
        sElementValue = yBuff
    End Sub

    Private Function lReadProperty(ByVal v_sNodeName As String, ByVal sElementValue As Integer) As Integer
        If mp_bSupportOptional = False Then
            Return ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText
        Else
            If ParentElement() Is Nothing Then
                Return sElementValue
            End If
            If ParentElement.GetElementsByTagName(v_sNodeName).Count > 0 Then
                Return ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText
            Else
                Return sElementValue
            End If
        End If
    End Function

    Private Function iReadProperty(ByVal v_sNodeName As String, ByVal sElementValue As Short) As Short
        If mp_bSupportOptional = False Then
            Return ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText
        Else
            If ParentElement() Is Nothing Then
                Return sElementValue
            End If
            If ParentElement.GetElementsByTagName(v_sNodeName).Count > 0 Then
                Return ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText
            Else
                Return sElementValue
            End If
        End If
    End Function

    Private Function sReadProperty(ByVal v_sNodeName As String, ByVal sElementValue As String) As String
        If mp_bSupportOptional = False Then
            Return ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText
        Else
            If ParentElement() Is Nothing Then
                Return sElementValue
            End If
            If ParentElement.GetElementsByTagName(v_sNodeName).Count > 0 Then
                If ParentElement.GetElementsByTagName(v_sNodeName).Item(0).ParentNode.Name = ParentElement().Name Then
                    Dim sReturn As String
                    sReturn = ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText
                    Return sReturn
                Else
                    Return sElementValue
                End If
            Else
                Return sElementValue
            End If
        End If
    End Function

    Private Function bReadProperty(ByVal v_sNodeName As String, ByVal bElementValue As Boolean) As Boolean
        If mp_bSupportOptional = False Then
            If ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText = "false" Or ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText = "0" Then
                Return False
            Else
                Return True
            End If
        Else
            If ParentElement() Is Nothing Then
                Return bElementValue
            End If
            If ParentElement.GetElementsByTagName(v_sNodeName).Count > 0 Then
                If ParentElement.GetElementsByTagName(v_sNodeName).Item(0).ParentNode.Name = ParentElement().Name Then
                    If ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText = "false" Or ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText = "0" Then
                        Return False
                    Else
                        Return True
                    End If
                Else
                    Return bElementValue
                End If
            Else
                Return bElementValue
            End If
        End If
    End Function

    Private Function dtReadProperty(ByVal v_sNodeName As String, ByVal dtElementValue As System.DateTime) As System.DateTime
        If mp_bSupportOptional = False Then
            Return mp_dtGetDateFromXML(ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText)
        Else
            If ParentElement() Is Nothing Then
                Return dtElementValue
            End If
            If ParentElement.GetElementsByTagName(v_sNodeName).Count > 0 Then
                Return mp_dtGetDateFromXML(ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText)
            Else
                Return dtElementValue
            End If
        End If
    End Function

    Private Function mp_dtGetDateFromXML(ByVal sParam As String) As System.DateTime
        Dim dtReturn As System.DateTime
        Dim lYear As Integer = System.Convert.ToInt32(sParam.Substring(0, 4))
        Dim lMonth As Integer = System.Convert.ToInt32(sParam.Substring(5, 2))
        Dim lDay As Integer = System.Convert.ToInt32(sParam.Substring(8, 2))
        Dim lHours As Integer = System.Convert.ToInt32(sParam.Substring(11, 2))
        Dim lMinutes As Integer = System.Convert.ToInt32(sParam.Substring(14, 2))
        Dim lSeconds As Integer = System.Convert.ToInt32(sParam.Substring(17, 2))
        dtReturn = New System.DateTime(lYear, lMonth, lDay, lHours, lMinutes, lSeconds)
        Return dtReturn
    End Function

    Private Function fReadProperty(ByVal v_sNodeName As String, ByVal fElementValue As Single) As Single
        If mp_bSupportOptional = False Then
            Return ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText
        Else
            If ParentElement() Is Nothing Then
                Return fElementValue
            End If
            If ParentElement.GetElementsByTagName(v_sNodeName).Count > 0 Then
                Return ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText
            Else
                Return fElementValue
            End If
        End If
    End Function

    Private Function yReadProperty(ByVal v_sNodeName As String, ByVal yElementValue As Integer) As Integer
        If mp_bSupportOptional = False Then
            Return ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText
        Else
            If ParentElement() Is Nothing Then
                Return yElementValue
            End If
            If ParentElement.GetElementsByTagName(v_sNodeName).Count > 0 Then
                Return ParentElement.GetElementsByTagName(v_sNodeName).Item(0).InnerText
            Else
                Return yElementValue
            End If
        End If

    End Function

    Friend Sub WriteObject(ByVal sObjectText As String)
        Dim xDoc1 As System.Xml.XmlDocument
        Dim oNodeBuff As System.Xml.XmlElement
        xDoc1 = New System.Xml.XmlDocument()
        xDoc1.LoadXml(sObjectText)
        oNodeBuff = xDoc.ImportNode(xDoc1.DocumentElement, True)
        ParentElement.AppendChild(oNodeBuff)
    End Sub

    Friend Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Object)
        Dim oNodeBuff As System.Xml.XmlElement
        oNodeBuff = xDoc.CreateElement(sElementName)
        If sElementValue.GetType().FullName = "System.DateTime" Then
            oNodeBuff.InnerText = mp_sGetXMLDateString(sElementValue)
        ElseIf sElementValue.GetType().FullName = "System.Boolean" Then
            If System.Convert.ToBoolean(sElementValue) = True Then
                If mp_bBoolsAreNumeric = False Then
                    oNodeBuff.InnerText = "true"
                Else
                    oNodeBuff.InnerText = "1"
                End If
            Else
                If mp_bBoolsAreNumeric = False Then
                    oNodeBuff.InnerText = "false"
                Else
                    oNodeBuff.InnerText = "0"
                End If
            End If
        ElseIf sElementValue.GetType().FullName = "System.String" Then
            oNodeBuff.InnerText = sElementValue
        Else
            oNodeBuff.InnerText = sElementValue.ToString()
        End If
        ParentElement.AppendChild(oNodeBuff)
    End Sub

    Private Function mp_sGetXMLDateString(ByVal dtParam As System.DateTime) As String
        Return g_Format(dtParam.Year, "0000") & "-" & g_Format(dtParam.Month, "00") & "-" & g_Format(dtParam.Day, "00") & "T" & g_Format(dtParam.Hour, "00") & ":" & g_Format(dtParam.Minute, "00") & ":" & g_Format(dtParam.Second, "00")
    End Function

    '// Microsoft Project Integration

    Friend Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Time)
        Dim oNodeBuff As System.Xml.XmlElement
        oNodeBuff = xDoc.CreateElement(sElementName)
        oNodeBuff.InnerText = sElementValue.ToString()
        ParentElement.AppendChild(oNodeBuff)
    End Sub

    Friend Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Duration)
        Dim oNodeBuff As System.Xml.XmlElement
        oNodeBuff = xDoc.CreateElement(sElementName)
        oNodeBuff.InnerText = sElementValue.ToString()
        ParentElement.AppendChild(oNodeBuff)
    End Sub

    Friend Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Decimal)
        Dim oNodeBuff As System.Xml.XmlElement
        oNodeBuff = xDoc.CreateElement(sElementName)
        oNodeBuff.InnerText = sElementValue.ToString()
        ParentElement.AppendChild(oNodeBuff)
    End Sub

    Friend Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Single)
        Dim oNodeBuff As System.Xml.XmlElement
        oNodeBuff = xDoc.CreateElement(sElementName)
        Dim sReturn As String = sElementValue.ToString()
        If InStr(sReturn, "E") <> 0 Then
            Dim oDecimal As Decimal = System.Convert.ToDecimal(sElementValue)
            sReturn = oDecimal.ToString()
        End If
        oNodeBuff.InnerText = sReturn
        ParentElement.AppendChild(oNodeBuff)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As Time)
        If mp_bSupportOptional = False Then
            sElementValue.FromString(ParentElement.GetElementsByTagName(sElementName).Item(0).InnerText)
        Else
            If ParentElement() Is Nothing Then
                Return
            End If
            If ParentElement.GetElementsByTagName(sElementName).Count > 0 Then
                sElementValue.FromString(ParentElement.GetElementsByTagName(sElementName).Item(0).InnerText)
            End If
        End If

    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As Duration)
        If mp_bSupportOptional = False Then
            sElementValue.FromString(ParentElement.GetElementsByTagName(sElementName).Item(0).InnerText)
        Else
            If ParentElement() Is Nothing Then
                Return
            End If
            If ParentElement.GetElementsByTagName(sElementName).Count > 0 Then
                sElementValue.FromString(ParentElement.GetElementsByTagName(sElementName).Item(0).InnerText)
            End If
        End If

    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As Decimal)
        If mp_bSupportOptional = False Then
            sElementValue = System.Convert.ToDecimal(ParentElement.GetElementsByTagName(sElementName).Item(0).InnerText)
        Else
            If ParentElement() Is Nothing Then
                Return
            End If
            If ParentElement.GetElementsByTagName(sElementName).Count > 0 Then
                sElementValue = System.Convert.ToDecimal(ParentElement.GetElementsByTagName(sElementName).Item(0).InnerText)
            End If
        End If
    End Sub



End Class
