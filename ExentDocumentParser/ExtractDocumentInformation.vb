Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Text.RegularExpressions

Module ExtractDocumentInformation
    Enum FieldTypes
        [String] = 0
        Currency = 1
        [Date] = 2
        Quantity = 3
        ABN = 4
    End Enum

    Dim ExcelApp As Excel.Application
    Dim f As FormActivity
    Dim SheetContents As Object
    Dim ShowOfficeApplications As Boolean

    Public Sub GetDocumentContents(WordDocument As String, ExcelWorkbook As String, ProcessingTemplatesFolder As String, ByRef HeaderData As Dictionary(Of String, String), ByRef DetailsData As DataTable, Optional DisplayOfficeApps As Boolean = False)
        Dim wb As Excel.Workbook
        Dim ws As Excel.Worksheet
        Dim ProcessingData As Data.DataTable
        Dim ProcessingTemplate As String
        Dim ABN As String
        Dim Value As String = ""
        Dim InvoiceTotal As Decimal

        Dim CompanyName As String = ""

        f = New FormActivity
        f.Show()

        ' ***************************************************************************
        UpdateStatus(f, "Loading Excel...")
        ExcelApp = New Excel.Application
        ExcelApp.Visible = DisplayOfficeApps

        ' See if the Excel Workbook already exists (all converted - most likely only going to ne used in testing)
        If File.Exists(ExcelWorkbook) Then
            wb = ExcelApp.Workbooks.Open(ExcelWorkbook)
        ElseIf WordDocument <> "" Then
            ' Create New Excel Document 
            wb = ExcelApp.Workbooks.Add
            wb.SaveAs(ExcelWorkbook)

            UpdateStatus(f, "Converting Content to Spreadsheet...")
            ConvertDOCXToXLSX(f, WordDocument, ExcelApp, wb, DisplayOfficeApps)
        Else
            ' Nothing to Do...

            Exit Sub
        End If

        ' Set the Worksheet
        ws = wb.ActiveSheet

        'UnMergeCells(ws)
        'DeleteEmptyRows(ws)

        ' ***************************************************************************
        Dim Myrange As Excel.Range = ws.Range(ws.Cells(1, 1), ws.Cells(ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count))

        ' Get the Sheet Contents into Array for fast reading
        SheetContents = Myrange.Value

        ' Load the Data from the Default Template
        UpdateStatus(f, "Loading the Default Processing Template...")
        ProcessingTemplate = ProcessingTemplatesFolder + "Default.xlsx"
        ProcessingData = GetProcessingTemplateData(ProcessingTemplate)

        Log(LogLevels.Trace, "Setting Defaults for Invoice...")
        AddHeaderData(HeaderData, "InvoiceValid", "TRUE")
        AddHeaderData(HeaderData, "Exceptions", "")

        ' ***************************************************************************
        ' **** Get ABN **** REQUIRED ****
        Log(LogLevels.Trace, "Getting ABN...")
        ABN = GetValueFromSpreadsheet("ABN", FieldTypes.ABN, ProcessingData)
        ' Remnove Spaces from ABN
        ABN = ABN.Replace(" ", "")
        AddHeaderData(HeaderData, "ABN", ABN)
        If ABN = "" Then SetInvoiceInvalid(HeaderData, "Missing ABN!")
        Log(LogLevels.Trace, "ABN = " + ABN)

        ' **** Get Company Name ****
        ' Use Web Service
        If ABN <> "" Then
            Log(LogLevels.Trace, "Getting Company Name...")
            Value = GetCompanyName(ABN)
            Log(LogLevels.Trace, "Company Name = " + Value)
        Else
            Value = ""
        End If
        AddHeaderData(HeaderData, "InvoiceCompany", Value)

        ' ***************************************************************************
        ' **** Check for Custom Processing Template ****
        Log(LogLevels.Trace, "Checking for Processing Templates Folder...")
        If Directory.GetFiles(ProcessingTemplatesFolder, ABN + "*.xlsx").Count > 0 Then
            UpdateStatus(f, "Loading the Custom Processing Template...")
            ' Get the first file in the list (should only be 1)
            ProcessingTemplate = Directory.GetFiles(ProcessingTemplatesFolder, ABN + "*.xlsx")(0)
            ' Load the Custom Processing Template
            ProcessingData = GetProcessingTemplateData(ProcessingTemplate)
        End If

        ' ***************************************************************************
        UpdateStatus(f, "Closing Excel...")
        wb.Close(True)
        ExcelApp.Quit()

        releaseObject(ExcelApp)
        releaseObject(wb)

        ' ***************************************************************************
        ' **** Get Invoice Number **** REQUIRED ****
        Log(LogLevels.Trace, "Getting Invoice Number...")
        Value = GetValueFromSpreadsheet("InvoiceNumber", FieldTypes.String, ProcessingData)
        Log(LogLevels.Trace, "Invoice Number = " + Value)
        AddHeaderData(HeaderData, "InvoiceNumber", Value)
        If Value = "" Then SetInvoiceInvalid(HeaderData, "Missing Invoice Number!")

        ' **** Get Order Number **** REQUIRED ****
        Value = GetValueFromSpreadsheet("OrderNumber", FieldTypes.String, ProcessingData)
        AddHeaderData(HeaderData, "OrderNumber", Value)
        If Value = "" Then SetInvoiceInvalid(HeaderData, "Missing Order Number!")

        ' **** Get Date ****
        Value = GetValueFromSpreadsheet("InvoiceDate", FieldTypes.Date, ProcessingData)
        ' Convert to Standard Format
        If IsDate(Value) Then
            Value = CDate(Value).ToString("d/MM/yyyy")
        Else
            Value = ""
        End If
        AddHeaderData(HeaderData, "InvoiceDate", Value)

        ' **** Get Due Date ****
        Value = GetValueFromSpreadsheet("InvoiceDueDate", FieldTypes.Date, ProcessingData)
        ' Convert to Standard Format
        If IsDate(Value) Then
            Value = CDate(Value).ToString("d/MM/yyyy")
        Else
            Value = ""
        End If
        AddHeaderData(HeaderData, "InvoiceDueDate", Value)

        ' **** Get Terms ****
        Value = GetValueFromSpreadsheet("InvoiceTerms", FieldTypes.String, ProcessingData)
        ' Convert to Standard Format
        If Value <> "" Then
            ' Get numerical numbers only (check if between 1 and 100 - 7/14/30/60/90)
            Dim TermsDays As String = Regex.Replace(Value, "[^0-9 ]", "")
            ' Get the first 3 characters only - in case other numbers in string
            TermsDays = Strings.Left(TermsDays.Trim, 3).Trim
            ' Get the Text Only to find if Days/Invoice/Statement 
            Dim TermsPeriod As String = Regex.Replace(Value, "[^a-zA-Z ]", "")

            If TermsPeriod.ToUpper.Contains("DAY") Or IsNumeric(TermsDays) Then Value = TermsDays + "DAYS"
            If TermsPeriod.ToUpper.Contains("INV") Or TermsPeriod.ToUpper.Contains("NET") Then Value = Value + " INVOICE"
            If TermsPeriod.ToUpper.Contains("MONTH") Or TermsPeriod.ToUpper.Contains("EOM") Then Value = Value + " EOM"
            If TermsPeriod.ToUpper.Contains("ST") Then Value = Value + " STATEMENT"
        End If
        AddHeaderData(HeaderData, "InvoiceTerms", Value)

        ' **** Get Total Value of Invoice **** REQUIRED ****
        Value = GetValueFromSpreadsheet("InvoiceTotal", FieldTypes.Currency, ProcessingData)
        AddHeaderData(HeaderData, "InvoiceTotal", Value)
        If IsNumeric(Value) Then
            InvoiceTotal = CDec(Value)
        Else
            SetInvoiceInvalid(HeaderData, "Missing Invoice Total!")
        End If

        ' **** Get Total Value of GST ****
        Value = GetValueFromSpreadsheet("InvoiceGST", FieldTypes.Currency, ProcessingData)
        AddHeaderData(HeaderData, "InvoiceGST", Value)

        ' **** Get Invoice Details ****
        Dim HeaderRow As Long = 0       ' Row Number that Header Starts
        Dim HeaderRowSpan As Long = 1   ' How many rows in Header
        Dim DetailsRowSpan As Long = 1         ' How many rows in Detail Lines
        Dim LastRow As Long = 0         ' Last Row of Details

        DetailsData = CreasteDetailsDataTable()

        ' Get the Start Row of Details
        HeaderRow = GetRowFromSpreadsheet("InvoiceHeader", FieldTypes.String, ProcessingData, HeaderRowSpan)

        If HeaderRow > 0 Then

            ' Get the End Row of Details
            LastRow = GetRowFromSpreadsheet("InvoiceDetailsEnd", FieldTypes.String, ProcessingData) - 1

                ' Get the Row Span of Details
                GetRowFromSpreadsheet("InvoiceDetails", FieldTypes.String, ProcessingData, DetailsRowSpan)

                If LastRow > 0 Then
                    ' Add Records for each row
                    Dim RowID As Integer = 1
                    For DetailRow As Long = HeaderRow + HeaderRowSpan To LastRow Step DetailsRowSpan
                        'Dim dr As DataRow = DetailsData.NewRow
                        'dr("ID") = ID
                        DetailsData.Rows.Add(RowID)
                        RowID += 1
                    Next

                    For DetailsCol As Long = 1 To 7
                        Dim DetailHeader = ""
                        Dim DetailColumn As Long
                        Dim CellColumnSplit As String = ""
                        Dim CellColumn As String = ""

                        Select Case DetailsCol
                            Case 1 : DetailHeader = "ItemsCode"
                            Case 2 : DetailHeader = "ItemsDescription"
                            Case 3 : DetailHeader = "ItemsQty"
                            Case 4 : DetailHeader = "ItemsExTax"
                            Case 5 : DetailHeader = "ItemsIncTax"
                            Case 6 : DetailHeader = "ItemsExExtended"
                            Case 7 : DetailHeader = "ItemsIncExtended"
                        End Select

                        ' Find the header in the header cells and get the column it is in
                        DetailColumn = GetColForDetailFromSpreadsheet(ProcessingData, DetailHeader, HeaderRow, LastRow, CellColumnSplit, CellColumn)

                        If DetailColumn > 0 Then
                            ' Go through each of the rows below and update the Details Table
                            RowID = 1
                            For DetailRow As Long = HeaderRow + HeaderRowSpan To LastRow Step DetailsRowSpan
                                Dim DetailValue As String = SheetContents(DetailRow, DetailColumn)

                                If CellColumnSplit <> "" And IsNumeric(CellColumn) Then
                                    Dim CellsColumns() As String = Split(DetailValue, CellColumnSplit.Replace("""", ""))
                                    ' Remove any empty values from array
                                    For i = CellsColumns.Length - 1 To 0 Step -1
                                        If CellsColumns(i).Trim = "" Then
                                            CellsColumns = CellsColumns.Where(Function(item, index) index <> i).ToArray
                                        Else
                                            CellsColumns(i) = CellsColumns(i).Trim
                                        End If
                                    Next

                                    If CellsColumns.Count >= CLng(CellColumn) Then
                                        DetailValue = CellsColumns(CLng(CellColumn) - 1) ' -1 is due to 0 index array
                                    End If
                                End If

                                Select Case DetailsCol
                                    Case 1 : DetailsData.Rows(RowID - 1)("Code") = DetailValue
                                    Case 2 : DetailsData.Rows(RowID - 1)("Description") = DetailValue
                                    Case 3 : If IsNumeric(DetailValue) Then DetailsData.Rows(RowID - 1)("Qty") = CDec(DetailValue)
                                    Case 4 : If IsNumeric(DetailValue) Then DetailsData.Rows(RowID - 1)("UnitExTax") = CDec(DetailValue)
                                    Case 5 : If IsNumeric(DetailValue) Then DetailsData.Rows(RowID - 1)("UnitIncTax") = CDec(DetailValue)
                                    Case 6 : If IsNumeric(DetailValue) Then DetailsData.Rows(RowID - 1)("ExtendedExTax") = CDec(DetailValue)
                                    Case 7 : If IsNumeric(DetailValue) Then DetailsData.Rows(RowID - 1)("ExtendedIncTax") = CDec(DetailValue)
                                End Select

                                RowID += 1
                            Next
                        End If
                    Next
                Else
                    SetInvoiceInvalid(HeaderData, "Missing Invoice Details - Can not find end of Details!")
                End If
            Else
                SetInvoiceInvalid(HeaderData, "Missing Invoice Details - Can not find Header!")
        End If

        ' Validate the Detail Line Contents
        For RowIdx As Long = 0 To DetailsData.Rows.Count - 1
            If DetailsData.Rows(RowIdx)("Qty").ToString = "" Then
                SetInvoiceInvalid(HeaderData, "Missing Invoice Details - No Quantiy for Line " + (RowIdx + 1).ToString + "!")
            ElseIf DetailsData.Rows(RowIdx)("UnitExTax").ToString = "" And DetailsData.Rows(RowIdx)("UnitIncTax").ToString = "" And DetailsData.Rows(RowIdx)("ExtendedExTax").ToString = "" And DetailsData.Rows(RowIdx)("ExtendedIncTax").ToString = "" Then
                SetInvoiceInvalid(HeaderData, "Missing Invoice Details - No Pricing for Line " + (RowIdx + 1).ToString + "!")
            Else

                ' Get Invoice GST Rate (Default 10%) - TO DO - Should we validate this against the Doc Total and GDT Total
                Dim GSTRate As Decimal = 10

                ' Update Missing Data
                ' Missing Unit Ex Tax & Inc Tax but not ExtendedExTax
                If DetailsData.Rows(RowIdx)("UnitExTax").ToString = "" And DetailsData.Rows(RowIdx)("UnitIncTax").ToString = "" And DetailsData.Rows(RowIdx)("ExtendedExTax").ToString <> "" Then
                    DetailsData.Rows(RowIdx)("UnitExTax") = Math.Round(DetailsData.Rows(RowIdx)("ExtendedExTax") / DetailsData.Rows(RowIdx)("Qty"), 2)
                End If

                ' Missing Unit Ex Tax & Inc Tax but not ExtendedIncTax
                If DetailsData.Rows(RowIdx)("UnitExTax").ToString = "" And DetailsData.Rows(RowIdx)("UnitIncTax").ToString = "" And DetailsData.Rows(RowIdx)("ExtendedIncTax").ToString <> "" Then
                    DetailsData.Rows(RowIdx)("UnitIncTax") = Math.Round(DetailsData.Rows(RowIdx)("ExtendedIncTax") / DetailsData.Rows(RowIdx)("Qty"), 2)
                End If

                ' Missing Unit Ex Tax & Inc Tax but not ExtendedExTax
                If DetailsData.Rows(RowIdx)("ExtendedExTax").ToString = "" And DetailsData.Rows(RowIdx)("ExtendedIncTax").ToString = "" And DetailsData.Rows(RowIdx)("UnitExTax").ToString <> "" Then
                    DetailsData.Rows(RowIdx)("ExtendedExTax") = Math.Round(DetailsData.Rows(RowIdx)("UnitExTax") * DetailsData.Rows(RowIdx)("Qty"), 2)
                End If

                ' Missing Extended Ex Tax & Exetended Inc Tax but not Unit Inc Tax
                If DetailsData.Rows(RowIdx)("ExtendedExTax").ToString = "" And DetailsData.Rows(RowIdx)("ExtendedIncTax").ToString = "" And DetailsData.Rows(RowIdx)("UnitIncTax").ToString <> "" Then
                    DetailsData.Rows(RowIdx)("ExtendedIncTax") = Math.Round(DetailsData.Rows(RowIdx)("UnitIncTax") * DetailsData.Rows(RowIdx)("Qty"), 2)
                End If

                ' Missing Unit Ex Tax - ExtTax = IncTax / 1.1
                If DetailsData.Rows(RowIdx)("UnitExTax").ToString = "" And DetailsData.Rows(RowIdx)("UnitIncTax").ToString <> "" Then
                    DetailsData.Rows(RowIdx)("UnitExTax") = Math.Round(DetailsData.Rows(RowIdx)("UnitIncTax") / ((100 + GSTRate) / 100), 2)

                    ' Missing Unit Inc Tax - IncTax = ExtTax * 1.1
                ElseIf DetailsData.Rows(RowIdx)("UnitExTax").ToString <> "" And DetailsData.Rows(RowIdx)("UnitIncTax").ToString = "" Then
                    DetailsData.Rows(RowIdx)("UnitIncTax") = Math.Round(DetailsData.Rows(RowIdx)("UnitExTax") * ((100 + GSTRate) / 100), 2)
                End If

                ' Missing Exetended Ex Tax - ExtTax = IncTax / 1.1
                If DetailsData.Rows(RowIdx)("ExtendedExTax").ToString = "" And DetailsData.Rows(RowIdx)("ExtendedIncTax").ToString <> "" Then
                    DetailsData.Rows(RowIdx)("ExtendedExTax") = Math.Round(DetailsData.Rows(RowIdx)("ExtendedIncTax") / ((100 + GSTRate) / 100), 2)

                    ' Missing Unit Inc Tax - IncTax = ExtTax * 1.1
                ElseIf DetailsData.Rows(RowIdx)("ExtendedExTax").ToString <> "" And DetailsData.Rows(RowIdx)("ExtendedIncTax").ToString = "" Then
                    DetailsData.Rows(RowIdx)("ExtendedIncTax") = Math.Round(DetailsData.Rows(RowIdx)("ExtendedExTax") * ((100 + GSTRate) / 100), 2)
                End If
            End If
        Next

        f.Close()

    End Sub

    Private Sub UnMergeCells(ws As Excel.Worksheet)
        Dim C As Excel.Range

        ExcelApp.FindFormat.Clear()
        ExcelApp.FindFormat.MergeCells = True
        C = ws.Cells.Find("", SearchFormat:=True)
        Do While Not C Is Nothing
            C.UnMerge()
            C = ws.Cells.Find("", SearchFormat:=True)
        Loop
        ExcelApp.FindFormat.Clear()
    End Sub

    Private Sub DeleteEmptyRows(ws As Excel.Worksheet)
        ' Clear the Empty Cells
        ' For Each Row (backwards - to allow for deletes
        For row As Long = ws.UsedRange.Rows.Count To 1 Step -1
            Dim RowEmpty As Boolean = True
            ' Go through Columns
            For col As Long = 1 To ws.UsedRange.Columns.Count
                Dim cell As Excel.Range = ws.Cells(row, col)

                If String.IsNullOrEmpty(cell.Value) Then
                    RowEmpty = False
                End If
            Next

            If RowEmpty Then
                ws.Rows(row).EntireRow.Delete
            End If
        Next
    End Sub

    Private Sub DeleteEmptyRowsAndCells(ws As Excel.Worksheet)
        ' Clear the Empty Cells
        ' For Each Row (backwards - to allow for deletes
        For row As Long = ws.UsedRange.Rows.Count To 1 Step -1
            Dim RowEmpty As Boolean = True
            ' Go through Columns, If empty delete the cell
            For col As Long = ws.UsedRange.Columns.Count - 1 To 1 Step -1
                Dim cell As Excel.Range = ws.Cells(row, col)
                Dim DeleteCell As Boolean = False

                If Not String.IsNullOrEmpty(cell.Value) Then
                    If cell.Value.ToString = "" Then
                        DeleteCell = True
                    Else
                        RowEmpty = False
                    End If
                Else
                    DeleteCell = True
                End If

                If DeleteCell Then cell.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)
            Next

            If RowEmpty Then
                ws.Rows(row).EntireRow.Delete
            End If
        Next
    End Sub

    Private Sub AddHeaderData(d As Dictionary(Of String, String), Key As String, Data As String)
        If d.ContainsKey(Key) Then
            d(Key) = Data
        Else
            d.Add(Key, Data)
        End If

    End Sub

    Private Sub SetInvoiceInvalid(d As Dictionary(Of String, String), Exception As String)
        If d.ContainsKey("InvoiceValid") Then
            d("InvoiceValid") = "FALSE"
        Else
            d.Add("InvoiceValid", "FALSE")
        End If

        If d.ContainsKey("Exceptions") Then
            d("Exceptions") = d("Exceptions") + Environment.NewLine + Exception
        Else
            d.Add("Exceptions", Exception)
        End If

    End Sub

    Private Function GetProcessingTemplateData(ProcessingTemplate As String) As Data.DataTable
        Dim wbTemplate As Excel.Workbook
        Dim wsTemplate As Excel.Worksheet

        If IsNothing(ExcelApp) Then
            ExcelApp = New Excel.Application
            ExcelApp.Visible = True
        End If

        wbTemplate = ExcelApp.Workbooks.Open(ProcessingTemplate,, True)

        ' Set the Worksheet
        wsTemplate = wbTemplate.ActiveSheet

        Dim dt As New Data.DataTable

        ' Create four typed columns in the DataTable.
        dt.Columns.Add("Field", GetType(String))
        dt.Columns.Add("Comment", GetType(String))

        dt.Columns.Add("SearchText", GetType(String))
        dt.Columns.Add("SearchDirection", GetType(String))

        dt.Columns.Add("AnchorPosition", GetType(String))
        dt.Columns.Add("AnchorCellsToMove", GetType(String))
        dt.Columns.Add("AnchorRowStart", GetType(String))
        dt.Columns.Add("AnchorRowEnd", GetType(String))
        dt.Columns.Add("AnchorColumnStart", GetType(String))
        dt.Columns.Add("AnchorColumnEnd", GetType(String))

        dt.Columns.Add("ValueMaxChars", GetType(String))

        dt.Columns.Add("RowSpan", GetType(String))

        For DetailsCol As Long = 1 To 7
            Dim DetailHeader As String = ""

            Select Case DetailsCol
                Case 1 : DetailHeader = "ItemsCode"
                Case 2 : DetailHeader = "ItemsDescription"
                Case 3 : DetailHeader = "ItemsQty"
                Case 4 : DetailHeader = "ItemsExTax"
                Case 5 : DetailHeader = "ItemsIncTax"
                Case 6 : DetailHeader = "ItemsExExtended"
                Case 7 : DetailHeader = "ItemsIncExtended"
            End Select

            dt.Columns.Add(DetailHeader + "Header", GetType(String))
            dt.Columns.Add(DetailHeader + "SheetColumn", GetType(String))
            dt.Columns.Add(DetailHeader + "CellColumnSplit", GetType(String))
            dt.Columns.Add(DetailHeader + "CellColumn", GetType(String))
            dt.Columns.Add(DetailHeader + "Length", GetType(String))
        Next

        ' Get the Sheet Contents into Array for fast reading
        Dim SheetContents As Object
        SheetContents = wsTemplate.Range(wsTemplate.Cells(3, 1), wsTemplate.Cells(wsTemplate.UsedRange.Rows.Count, wsTemplate.UsedRange.Columns.Count)).Value

        ' Read Rows
        For r As Long = 1 To UBound(SheetContents)
            Dim dr As DataRow = dt.NewRow

            dr("Field") = SheetContents(r, 1)
            dr("Comment") = SheetContents(r, 2)

            dr("SearchText") = SheetContents(r, 3)
            dr("SearchDirection") = SheetContents(r, 4)

            dr("AnchorPosition") = SheetContents(r, 5)
            dr("AnchorCellsToMove") = SheetContents(r, 6)
            dr("AnchorRowStart") = SheetContents(r, 7)
            dr("AnchorRowEnd") = SheetContents(r, 8)
            dr("AnchorColumnStart") = SheetContents(r, 9)
            dr("AnchorColumnEnd") = SheetContents(r, 10)

            dr("ValueMaxChars") = SheetContents(r, 11)

            dr("RowSpan") = SheetContents(r, 12)

            ' Get the Detail Lines Settings
            Dim ColStart As Long = 13

            For DetailsCol As Long = 1 To 7
                Dim DetailHeader As String = ""

                Select Case DetailsCol
                    Case 1 : DetailHeader = "ItemsCode"
                    Case 2 : DetailHeader = "ItemsDescription"
                    Case 3 : DetailHeader = "ItemsQty"
                    Case 4 : DetailHeader = "ItemsExTax"
                    Case 5 : DetailHeader = "ItemsIncTax"
                    Case 6 : DetailHeader = "ItemsExExtended"
                    Case 7 : DetailHeader = "ItemsIncExtended"
                End Select

                If DetailsCol > 1 Then ColStart = ColStart + 5

                dr(DetailHeader + "Header") = SheetContents(r, ColStart)
                dr(DetailHeader + "SheetColumn") = SheetContents(r, ColStart + 1)
                dr(DetailHeader + "CellColumnSplit") = SheetContents(r, ColStart + 2)
                dr(DetailHeader + "CellColumn") = SheetContents(r, ColStart + 3)
                dr(DetailHeader + "Length") = SheetContents(r, ColStart + 4)
            Next

            dt.Rows.Add(dr)
        Next

        wbTemplate.Close(False)

        Return dt

    End Function

    Private Function CreasteDetailsDataTable() As Data.DataTable
        Dim dt As New Data.DataTable

        ' Create typed columns in the DataTable.
        dt.Columns.Add("ID", GetType(Integer))
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Description", GetType(String))
        dt.Columns.Add("Qty", GetType(Decimal))
        dt.Columns.Add("UnitExTax", GetType(Decimal))
        dt.Columns.Add("UnitIncTax", GetType(Decimal))
        dt.Columns.Add("ExtendedExTax", GetType(Decimal))
        dt.Columns.Add("ExtendedIncTax", GetType(Decimal))

        Return dt

    End Function

    Private Function GetCompanyName(ABN As String) As String
        Dim Search As httpXMLSearch
        Dim SearchPayload As String
        Dim MainName As String = ""
        Dim MainTradingName As String = ""
        Dim CompanyName As String = ""

        Log(LogLevels.Trace, "GetCompanyName - Declaring Search Variable...")
        Search = New httpXMLDocumentSearch

        Log(LogLevels.Trace, "GetCompanyName - Performing ABN Search...")
        SearchPayload = Search.ABNSearch(ABN, "n", "371f387c-0a3b-420b-ac39-00bb04f5b85f")

        Log(LogLevels.Trace, "GetCompanyName - Parsing Result...")
        Dim p As New XMLMessageParser(SearchPayload)

        Return p.GetCompanyName()

    End Function

    Private Function GetValueFromSpreadsheet(Field As String, FieldType As FieldTypes, TemplateData As Data.DataTable, Optional ByRef RowSpan As Long = 0) As String
        Dim sResult As String = ""
        Dim lRow As Long = 0

        ' Filter DataTable to Field
        Dim rows As Data.DataRow() = TemplateData.Select("Field = '" + Field + "'")

        For Each row As DataRow In rows
            GetContentForFieldUsingAnchor(Field, FieldType, row, sResult, lRow)
            If sResult <> "" Then
                ' For Details Header/Lines get if they span multiple Rows
                If IsNumeric(row("RowSpan").ToString) Then
                    RowSpan = row("RowSpan").ToString
                Else
                    RowSpan = 1
                End If

                Exit For
            End If
        Next

        Return sResult
    End Function

    Private Function GetRowFromSpreadsheet(Field As String, FieldType As FieldTypes, TemplateData As Data.DataTable, Optional ByRef RowSpan As Long = 0) As Long
        Dim sResult As String = ""
        Dim lRow As Long = 0

        ' Filter DataTable to Field
        Dim rows As Data.DataRow() = TemplateData.Select("Field = '" + Field + "'")

        For Each row As DataRow In rows
            GetContentForFieldUsingAnchor(Field, FieldType, row, sResult, lRow)

            If lRow > 0 Then
                ' For Details Header/Lines get if they span multiple Rows
                If IsNumeric(row("RowSpan").ToString) Then
                    RowSpan = row("RowSpan").ToString
                Else
                    RowSpan = 1
                End If

                Exit For
            End If
        Next

        Return lRow
    End Function

    Private Function GetColForDetailFromSpreadsheet(TemplateData As Data.DataTable, DetailHeader As String, RowStart As Long, RowEnd As Long, ByRef CellColumnSplit As String, ByRef CellColumn As String) As Long
        Dim sResult As String = ""
        Dim lRow As Long = 0
        Dim lCol As Long = 0

        ' Filter DataTable to Field
        Dim rows As Data.DataRow() = TemplateData.Select("Field = 'InvoiceDetails'")

        For Each row As DataRow In rows

            ' **** Find the Column holding the Data ****
            Dim MaxLength As String = row(DetailHeader + "Length").ToString

            ' If Column is specified in Template return this
            If IsNumeric(row(DetailHeader + "SheetColumn").ToString) Then
                lCol = CLng(row(DetailHeader + "SheetColumn").ToString)

            ElseIf row(DetailHeader + "Header").ToString <> "" Then ' Search for it in the Header Rows
                GetUsingAnchor(False, FieldTypes.String, row(DetailHeader + "Header").ToString, "", "", "", RowStart, RowEnd, "", "", MaxLength, sResult, lRow, lCol)
            End If

            ' Get the Cell Column Split Content
            CellColumnSplit = row(DetailHeader + "CellColumnSplit").ToString
            CellColumn = row(DetailHeader + "CellColumn").ToString
        Next

        Return lCol
    End Function

    Private Sub GetContentForFieldUsingAnchor(Field As String, FieldType As FieldTypes, row As Data.DataRow, ByRef ResultValue As String, Optional ByRef ResultRow As Long = 0, Optional ByRef ResultColumn As Long = 0)

        GetUsingAnchor(True,
                       FieldType,
                       row("SearchText").ToString,
                       row("SearchDirection").ToString,
                       row("AnchorPosition").ToString,
                       row("AnchorCellsToMove").ToString,
                       row("AnchorRowStart").ToString,
                       row("AnchorRowEnd").ToString,
                       row("AnchorColumnStart").ToString,
                       row("AnchorColumnEnd").ToString,
                       row("ValueMaxChars").ToString,
                       ResultValue,
                       ResultRow,
                       ResultColumn)

        If ResultValue <> "" Then
            Debug.Print("Found " + Field + " Value in Row " + (ResultRow).ToString + " Column " + (ResultColumn).ToString + " = " + ResultValue)
        End If

    End Sub

    Private Sub GetUsingAnchor(GetData As Boolean, FieldType As FieldTypes, SearchText As String, SearchDirection As String, AnchorPosition As String, AnchorCellsToMove As String, RowStart As String, RowEnd As String, ColStart As String, ColEnd As String, ValueMaxChars As String, ByRef ResultValue As String, Optional ByRef ResultRow As Long = 0, Optional ByRef ResultColumn As Long = 0)
        Dim sInput As String
        Dim sResult As String = ""
        Dim SearchPattern As String = ""
        Dim r As Long
        Dim c As Long
        Dim RowsAdded As Long
        Dim ColumnsAdded As Long

        ' Get the Defulat Row & Col Start, End & Step Settings (Direction - Top Down)
        Dim rStart As Long = 1
        Dim rEnd As Long = UBound(SheetContents, 1)
        Dim rStep As Long = 1
        Dim cStart As Long = 1
        Dim cEnd As Long = UBound(SheetContents, 2)
        Dim cStep As Long = 1

        ' Get the Search Direction
        If SearchDirection.ToUpper = "BOTTOM UP" Then
            rStart = UBound(SheetContents, 1)
            rEnd = 1
            rStep = -1
            cStart = UBound(SheetContents, 2)
            cEnd = 1
            cStep = -1
        End If
        ' Get the Search Direction
        If IsNumeric(RowStart) Then rStart = RowStart
        If IsNumeric(RowEnd) Then rEnd = RowEnd
        If IsNumeric(ColStart) Then cStart = ColStart
        If IsNumeric(ColEnd) Then cEnd = ColEnd

        For r = rStart To rEnd Step rStep
            For c = cStart To cEnd Step cStep
                RowsAdded = 0
                ColumnsAdded = 0

                SearchPattern = SearchText

                If SearchPattern <> "" Then
                    sInput = SheetContents(r, c)

                    If sInput <> "" Then

                        Dim regex As Regex = New Regex(SearchPattern, RegexOptions.IgnoreCase Or RegexOptions.Multiline)
                        Dim result As MatchCollection = regex.Matches(sInput)

                        If result.Count > 0 Then
                            If GetData Then
                                sResult = GetDataForField(sInput, SheetContents, AnchorPosition, AnchorCellsToMove, ValueMaxChars, result, FieldType, r, c, RowsAdded, ColumnsAdded)
                            Else
                                sResult = GetField(sInput, result)
                            End If

                            If sResult <> "" Then
                                ResultRow = r + RowsAdded
                                ResultColumn = c + ColumnsAdded
                            End If
                        End If
                    End If
                End If

                ' Check if we have found result - exit loop 
                If sResult <> "" Then Exit For
            Next

            ' Check if we have found result - exit loop 
            If sResult <> "" Then Exit For
        Next

        ResultValue = sResult

    End Sub

    Private Function GetDataForField(Input As String, SheetContents As Object, AnchorPosition As String, AnchorCellsToMove As String, ValueMaxChars As String, Result As MatchCollection, FieldType As FieldTypes, r As Long, c As Long, ByRef RowsAdded As Long, ByRef ColumnsAdded As Long) As String
        Dim sResult As String = ""
        Dim ResultLength As Long = 0
        If IsNumeric(ValueMaxChars) Then ResultLength = CLng(ValueMaxChars)

        ' Get the details after the search string
        ' Get the remaining text and trim
        sResult = Mid(Input, Result(0).Index + Result(0).Value.Length + 1).Trim

        ' If we have text after the identifier
        If InStr(sResult, vbCr) - 1 > Len(sResult) Then
            ' Remove everything past the eol
            sResult = Strings.Left(sResult, InStr(sResult, vbCr) - 1)
        End If

        ' Limit to Ascii Non Control Characters (i.e. remove tab, CR, LF, etc)
        sResult = Regex.Replace(sResult, "[^\x20-\x7E]", "")

        ' See if we have a specified number of cells to move
        Dim CellsToMove As Long = 0
        If IsNumeric(AnchorCellsToMove) Then CellsToMove = CLng(AnchorCellsToMove)

        Dim CellsToSearchRight As Integer = 1
        If AnchorPosition.ToUpper = "LEFT" Then
            CellsToSearchRight = 5
        End If

        ' If Empty or we need to search cells --> see if cells to right has value
        If sResult = "" Or CellsToSearchRight > 1 Then
            Dim StartCol As Long = 1

            If CellsToMove > 0 Then
                ColumnsAdded = CellsToMove
                sResult = SheetContents(r, c + ColumnsAdded)
                If sResult = "" Then
                    StartCol = CellsToMove + 1 ' Used in the case this value is blank we search the next column
                    CellsToSearchRight = CellsToMove + 5
                End If
            End If

            If sResult = "" Then
                ' Check Cells to right
                For ColumnsAdded = StartCol To CellsToSearchRight
                    If Not IsNothing(SheetContents(r, c + ColumnsAdded)) Then
                        sResult = SheetContents(r, c + ColumnsAdded).ToString
                        If sResult <> "" Then Exit For
                    End If
                Next
            End If
        End If

        Dim CellsToSearchDown As Integer = 1
        If AnchorPosition.ToUpper = "ABOVE" Then
            CellsToSearchDown = 5
        End If

        ' If Empty or we need to search cells --> see if cells below has value
        If sResult = "" Or CellsToSearchDown > 1 Then
            If CellsToMove > 0 Then
                RowsAdded = CellsToMove
                sResult = SheetContents(r + RowsAdded, c)
            Else
                ' Check Cell below
                For RowsAdded = 1 To CellsToSearchDown
                    If Not IsNothing(SheetContents(r + RowsAdded, c)) Then
                        sResult = SheetContents(r + RowsAdded, c).ToString
                        If sResult <> "" Then Exit For
                    Else ' Only Search Cells with values
                        RowsAdded = RowsAdded - 1
                    End If
                Next
            End If
        End If

        ' If we are specifying the number of characters, only get these
        If ResultLength > 0 Then
            sResult = Strings.Left(sResult, ResultLength)
        End If

        ' Trim Result
        sResult = sResult.Trim

        ' Limit to Ascii Non Control Characters (i.e. remove tab, CR, LF, etc)
        sResult = Regex.Replace(sResult, "[^\x20-\x7E]", "")

        ' Validate & Only Extract Required Data depending on Field Type
        Select Case FieldType
            Case FieldTypes.ABN  ' Only get numerical and decimal point characters and make sure it is 11 characters
                sResult = Regex.Replace(sResult, "[^0-9]", "").Trim
                If sResult.Length <> 11 Then sResult = ""

            Case FieldTypes.Currency  ' Only get numerical and decimal point characters
                sResult = Regex.Replace(sResult, "[^0-9.]", "").Trim

            Case FieldTypes.Date ' Only get numerical and - or / characters and Months 
                Dim regexValue As Regex = New Regex("\d{1,2}(/|-)(\d{1,2}|Jan|January|Feb|February|Mar|March|Apr|April|May|Jun|June|Jul|July|Aug|August|Sep|Septembet|Oct|October|Nov|November|Dec|December)\1(\d{4}|\d{2})", RegexOptions.IgnoreCase Or RegexOptions.Multiline)
                Result = regexValue.Matches(sResult)

                If Result.Count > 0 Then
                    sResult = Result(0).Value
                Else
                    sResult = ""
                End If

        End Select

        Return sResult
    End Function


    Private Function GetField(Input As String, Result As MatchCollection) As String
        Dim sResult As String = ""
        Dim ResultLength As Long = 0

        ' Get the details after the search string
        ' Get the remaining text and trim
        sResult = Input.Trim

        ' If we have text after the identifier
        If InStr(sResult, vbCr) - 1 > Len(sResult) Then
            ' Remove everything past the eol
            sResult = Strings.Left(sResult, InStr(sResult, vbCr) - 1)
        End If

        ' Limit to Ascii Non Control Characters (i.e. remove tab, CR, LF, etc)
        sResult = Regex.Replace(sResult, "[^\x20-\x7E]", "")

        Return sResult
    End Function

End Module
