Imports System.Text
Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO

Public Class CompanySearch_ResultsInterpreter
    '-----------------------------------------------------------------------------------------------
    ' Return payload as an XML String
    '-----------------------------------------------------------------------------------------------
    Public Function SerialisePayload(ByVal searchPayload As ABRSearch.Payload) As String
        Try
            Dim XmlStream As MemoryStream = New MemoryStream
            Dim XmlReader As StreamReader = New StreamReader(XmlStream)
            Dim Serializer As XmlSerializer = New XmlSerializer(GetType(ABRSearch.Payload))
            Serializer.Serialize(XmlStream, searchPayload)
            XmlStream.Seek(0, IO.SeekOrigin.Begin)
            Return XmlReader.ReadToEnd()
        Catch exp As Exception
            Throw
        End Try

    End Function

End Class
