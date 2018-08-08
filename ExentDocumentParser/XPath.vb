
Public Class XPath
   Public Enum Path
      commonRoot
      name
      abnValue
      state
      score
      postcode
      Exception
   End Enum
   ' Paths for SOAP messages
   Public Shared Function LocationSOAP(ByVal path As Path) As String
        Dim commonRoot As String = "//abn:ABRSearchByABNResponse/abn:ABRPayloadSearchResults/abn:response/abn:businessEntity"
        Select Case path
         Case path.commonRoot : Return commonRoot
            Case Path.name : Return "./abn:mainTradingName/abn:organisationName|./abn:mainName/abn:organisationName|./abn:legalName/abn:fullName|./abn:otherTradingName/abn:organisationName"
            Case path.abnValue : Return "./abn:ABN/abn:identifierValue"
         Case path.state : Return "./abn:mainBusinessPhysicalAddress/abn:stateCode"
         Case path.score : Return "./abn:legalName/abn:score|./abn:mainTradingName/abn:score|./abn:mainName/abn:score|./abn:otherTradingName/abn:score"
         Case path.postcode : Return "./abn:mainBusinessPhysicalAddress/abn:postcode"
         Case path.Exception : Return "abn:exception/abn:exceptionDescription"
      End Select
   End Function
End Class