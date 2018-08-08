Imports System.Net
Imports System.IO
'-----------------------------------------------------------------------------------------------
' Non-Strongly Typed Search - Builds the SOAP message as a string. 
'-----------------------------------------------------------------------------------------------
Public Class httpXMLDocumentSearch : Inherits httpXMLSearch
    '-----------------------------------------------------------------------------------------------
    ' Prefix in config file
    '-----------------------------------------------------------------------------------------------
    Protected Overrides ReadOnly Property StylePrefix() As String
        Get
            Return StyleEnum.Document.ToString()
        End Get
    End Property
    '-----------------------------------------------------------------------------------------------
    ' Return the SOAP message for a search by ABN using RPC style
    '-----------------------------------------------------------------------------------------------
    Protected Overrides Function BuildABNSOAPMessage(ByVal searchString As String,
                                           ByVal history As String,
                                           ByVal guid As String) As String


        Return "<?xml version=""1.0"" encoding=""utf-8""?>" &
                 "<soap:Envelope " &
                "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " &
                "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " &
                "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &
                "<soap:Body> " &
                "<ABRSearchByABN xmlns=""http://abr.business.gov.au/ABRXMLSearch/""> " &
                "<searchString>" & searchString & "</searchString>" &
                "<includeHistoricalDetails>" & history & "</includeHistoricalDetails>" &
                "<authenticationGuid>" & guid & "</authenticationGuid>" &
                "</ABRSearchByABN>" &
                "</soap:Body>" &
                "</soap:Envelope>"

    End Function
    '-----------------------------------------------------------------------------------------------
    ' Return the SOAP message for a search by ASIC
    '-----------------------------------------------------------------------------------------------
    Protected Overrides Function BuildASICSOAPMessage(ByVal searchString As String,
                                         ByVal history As String,
                                         ByVal guid As String) As String

        Return "<?xml version=""1.0"" encoding=""utf-8""?>" &
                "<soap:Envelope " &
                "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " &
                "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " &
                "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &
                "<soap:Body> " &
                "<ABRSearchByASIC xmlns=""http://abr.business.gov.au/ABRXMLSearch/""> " &
                "<searchString>" & searchString & "</searchString>" &
                "<includeHistoricalDetails>" & history & "</includeHistoricalDetails>" &
                "<authenticationGuid>" & guid & "</authenticationGuid>" &
                "</ABRSearchByASIC>" &
                "</soap:Body>" &
                "</soap:Envelope>"
    End Function
    '-----------------------------------------------------------------------------------------------
    ' Return the SOAP message for a search by Name
    '-----------------------------------------------------------------------------------------------
    Protected Overrides Function BuildNameSOAPMessage(ByVal searchString As String,
                                         ByVal ACT As String,
                                         ByVal NSW As String,
                                         ByVal NT As String,
                                         ByVal QLD As String,
                                         ByVal TAS As String,
                                         ByVal VIC As String,
                                         ByVal WA As String,
                                         ByVal SA As String,
                                         ByVal postcode As String,
                                         ByVal legalName As String,
                                         ByVal tradingName As String,
                                         ByVal guid As String) As String

        Return "<?xml version=""1.0"" encoding=""utf-8""?>" &
                "<soap:Envelope " &
                "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " &
                "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " &
                "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &
                "<soap:Body> " &
                "<ABRSearchByName xmlns=""http://abr.business.gov.au/ABRXMLSearch/""> " &
                "<externalNameSearch>" &
                "<name>" & EncodeXML(searchString) & "</name>" &
                "<filters>" &
                "<stateCode>" &
                "<ACT>" & ACT & "</ACT>" &
                "<NSW>" & NSW & "</NSW>" &
                "<NT>" & NT & "</NT>" &
                "<QLD>" & QLD & "</QLD>" &
                "<TAS>" & TAS & "</TAS>" &
                "<VIC>" & VIC & "</VIC>" &
                "<WA>" & WA & "</WA>" &
                "<SA>" & SA & "</SA>" &
                "</stateCode>" &
                "<postcode>" & postcode & "</postcode>" &
                "<nameType>" &
                "<legalName>" & legalName & "</legalName>" &
                "<tradingName>" & tradingName & "</tradingName>" &
                "</nameType>" &
                "</filters>" &
                "</externalNameSearch>" &
                "<authenticationGuid>" & guid & "</authenticationGuid>" &
                "</ABRSearchByName>" &
                "</soap:Body>" &
                "</soap:Envelope>"
    End Function
    '-----------------------------------------------------------------------------------------------
    ' Encodes a string as XML
    '-----------------------------------------------------------------------------------------------
    Private Function EncodeXML(ByVal searchString As String) As String
        Dim EncodedXML As String
        Dim Writer As New StringWriter
        Dim XMLWriter As New Xml.XmlTextWriter(Writer)
        XMLWriter.WriteString(searchString)
        Dim Reader As New StringReader(Writer.ToString)
        EncodedXML = Reader.ReadToEnd
        XMLWriter.Close()
        Writer.Close()
        Return EncodedXML
    End Function
End Class
