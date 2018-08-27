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
    Dim ShowOfficeApplications As Boolean

    Dim AppDataPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Exent\DocumentParser\")
    Dim WorkingPath As String = Path.Combine(AppDataPath, "Working")

    Public Sub GetDocumentContents(PDFDocument As String, ExcelWorkbook As String, WordDocument As String, ByRef HeaderData As Dictionary(Of String, String), ByRef DetailsData As DataTable, Optional DisplayOfficeApps As Boolean = False, Optional LogTextBox As Object = Nothing)
        Dim wb As Excel.Workbook
        Dim ws As Excel.Worksheet
        Dim PDFWorkingFile As String = ""
        'Dim ProcessingData As Data.DataTable

        TextBoxForLogging = LogTextBox

        f = New FormActivity
        f.Show()

        ' ***************************************************************************
        ' Make sure App Data Working Path exists
        If Not Directory.Exists(WorkingPath) Then Directory.CreateDirectory(WorkingPath)
        'Make sure working path is empty
        Dim file As String = ""
        Try
            For Each file In Directory.GetFiles(WorkingPath)
                IO.File.Delete(file)
            Next
        Catch ex As Exception
            Log(LogLevels.Warning, "Can't delete file - " + file)
        End Try

        ' ***************************************************************************
        ' Kill all Excel Processes
        For Each P As Process In System.Diagnostics.Process.GetProcessesByName("excel")
            P.Kill()
        Next

        ' ***************************************************************************
        ' Load Excel Application
        UpdateStatus(f, "Loading Excel...")
        ExcelApp = New Excel.Application
        ExcelApp.Visible = DisplayOfficeApps

        ' ***************************************************************************
        AddHeaderData(HeaderData, "DocumentValid", "TRUE")
        AddHeaderData(HeaderData, "Exceptions", "")

        If PDFDocument.Length > 0 Then
            UpdateStatus(f, "Converting PDF...")
            PDFWorkingFile = Path.Combine(WorkingPath, Path.GetFileName(PDFDocument))
            ' Move file to Working Folder
            IO.File.Copy(PDFDocument, PDFWorkingFile, True)
            PDFDocument = PDFWorkingFile

        ElseIf ExcelWorkbook.Length > 0 Then
            UpdateStatus(f, "Converting Excel Workbook...")

            SetInvoiceInvalid(HeaderData, "Excel Conversion Not Currently Supported - Please contact Exent Support.")

            PDFWorkingFile = Path.Combine(WorkingPath, Path.GetFileNameWithoutExtension(ExcelWorkbook), ".pdf")

            ' TO DO - Open using Excel and Export to PDF

        ElseIf WordDocument.Length > 0 Then
            UpdateStatus(f, "Converting Word Document...")

            SetInvoiceInvalid(HeaderData, "Word Conversion Not Currently Supported - Please contact Exent Support.")

            PDFWorkingFile = Path.Combine(WorkingPath, Path.GetFileNameWithoutExtension(WordDocument), ".pdf")

            ' TO DO - Open using Word and Export to PDF
        End If

        ' ***************************************************************************
        ' If PDF Document Exists - Extract Information
        If PDFWorkingFile.Length > 0 Then

            ' **** Create a new Temp Workbook in Excel and Add PDF File as OLE Object to find the Page Size of the PDF
            UpdateStatus(f, "Getting PDF Page Orientation...")
            wb = ExcelApp.Workbooks.Add()
            ws = wb.Worksheets(1)

            ws.OLEObjects.Add(Filename:=PDFWorkingFile, Link:=False, DisplayAsIcon:=False).Select

            Dim PDFHeight As Long = ws.OLEObjects(1).Height
            Dim PDFWidth As Long = ws.OLEObjects(1).Width
            Dim IsPDFPortrait As Boolean = PDFHeight > PDFWidth

            wb.Close(False)


            ' *** Classify PDF Document - Get the Document Type and Source Information
            UpdateStatus(f, "Classifying PDF Document...")
            ' Convert contents of PDF Invoice using PDF2XL with Full Document Layout for Orientation to Excel
            Dim ExcelWorkingFile As String = Path.Combine(WorkingPath, "FileContents.xlsx")
            Dim DocumentType As String = ""
            Dim DocumentSourceReference As String = ""

            ' Get the Document Classification Matrix & Search for content
            If ClassifyDocument(HeaderData, PDFWorkingFile, ExcelWorkingFile, IsPDFPortrait, DocumentType, DocumentSourceReference) Then
                Log(LogLevels.Info, "Document Type - " + DocumentType)
                Log(LogLevels.Info, "Document Source Reference - " + DocumentSourceReference)

                AddHeaderData(HeaderData, "DocumentType", DocumentType)
                AddHeaderData(HeaderData, "DocumentSourceReference", DocumentSourceReference)

                If DocumentType.Length = 0 Or DocumentSourceReference.Length = 0 Then
                    SetInvoiceInvalid(HeaderData, "Document Type/Source Reference Not Found in Document Classification Matrix.")
                Else
                    ' *** Read PDF Document
                    UpdateStatus(f, "Reading Contents of PDF using Document/Source Layout(s)...")
                    ' Convert contents of PDF Invoice using PDF2XL with the Specified Layout(s) for Document Type & Source to Excel
                    If ConvertPDFUsingLayoutFiles(HeaderData, PDFWorkingFile, ExcelWorkingFile, DocumentType, DocumentSourceReference) Then

                        ' *** Read Content From Spreadsheet
                        Dim ExtraTables As New Dictionary(Of String, DataTable)
                        ReadContentFromSpreadsheet(HeaderData, DetailsData, ExtraTables, ExcelWorkingFile)

                        ' *** Run Post Conversion Tasks
                        ' Open the tasks spreadsheet for the Document Type & Source 
                        Dim PostTasks As DataTable = GetPostConversionTasks(Path.Combine(AppDataPath, "Processing - Post Convestion Tasks\" + DocumentType + "_" + DocumentSourceReference + ".xlsx"))

                        ' Apply each of the tasks to the DataTables
                        For Each dr As DataRow In PostTasks.Rows
                            Select Case dr("Action").ToString.ToUpper
                                Case "Clean Invoice Details".ToUpper
                                    CleanInvoiceDetails(DetailsData, dr)

                                Case "Get Header Content".ToUpper, "Replace Header Content".ToUpper
                                    UpdateHeaderContent(HeaderData, DetailsData, ExtraTables, dr)

                            End Select
                        Next

                        ' *** Update Detail Values 
                        UpdateDetails(DetailsData)

                        ' *** Update and Validate Values 
                        UpdateAndValidateHeader(HeaderData, DetailsData)
                    End If
                End If
            End If
        End If

        UpdateStatus(f, "Closing Excel...")
        ExcelApp.Quit()
        releaseObject(ExcelApp)

        f.Close()

    End Sub

    Private Function ClassifyDocument(ByRef HeaderData As Dictionary(Of String, String), PDFWorkingFile As String, ExcelWorkingFile As String, IsPDFPortrait As Boolean, ByRef DocumentType As String, ByRef DocumentSourceReference As String) As Boolean
        Dim wb As Excel.Workbook
        Dim ws As Excel.Worksheet
        Dim LayoutFile As String

        If IsPDFPortrait Then
            LayoutFile = Path.Combine(AppDataPath, "Processing - Layouts", "AllContent_Portrait.layoutx")
        Else
            LayoutFile = Path.Combine(AppDataPath, "Processing - Layouts", "AllContent_Landscape.layoutx")
        End If

        Dim PDF2XLArguments As String = "-input=""" + PDFWorkingFile + """ " +
                                            "-layout=""" + LayoutFile + """ " +
                                            "-format=excelfile " +
                                            "-output=""" + ExcelWorkingFile + """ " +
                                            "-existingfile=replace " +
                                            "-noui"

        Dim startInfo As New ProcessStartInfo
        startInfo.FileName = My.Settings.PDF2XL_Executable
        startInfo.Arguments = PDF2XLArguments
        startInfo.UseShellExecute = True
        Dim Converter As System.Diagnostics.Process = Process.Start(startInfo)
        Dim timeout As Integer = 60000 '1 minute in milliseconds

        If Not Converter.WaitForExit(timeout) Then
            AddHeaderData(HeaderData, "DocumentValid", "FALSE")
            AddHeaderData(HeaderData, "Exceptions", "Timeout waiting for PDF2XL Conversion.")

            Return False
        Else
            ' Open Document Classification Matrix Spreadsheet and read into DataTable
            Dim Matrix As New DataTable
            Matrix = GetDocumentClassificationMatrix(Path.Combine(AppDataPath, "Processing - Document Classification\Document Classification Matrix.xlsx"))

            ' Open Excel Working File Spreadsheet
            wb = ExcelApp.Workbooks.Open(ExcelWorkingFile, [ReadOnly]:=True)

            ' Set the Worksheet
            ws = wb.Sheets(1)

            ' Get the Sheet Contents into Array for fast reading
            Dim Contents As Object
            Contents = ws.Range(ws.Cells(1, 1), ws.Cells(ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count)).Value

            ' Close Excel Files
            wb.Close(False)

            ' Iterate through DataTable looking for match
            For Each dr As DataRow In Matrix.Rows
                ' Search for Values
                Dim SearchText1 As String = dr("SearchText1").ToString
                Dim SearchText2 As String = dr("SearchText2").ToString
                Dim SearchText3 As String = dr("SearchText3").ToString

                Dim LineFound As Boolean
                LineFound = (RegExSearch(Contents, SearchText1) <> "")
                If LineFound And SearchText2 <> "" Then LineFound = (RegExSearch(Contents, SearchText2) <> "")
                If LineFound And SearchText3 <> "" Then LineFound = (RegExSearch(Contents, SearchText3) <> "")

                If LineFound Then
                    DocumentType = dr("DocumentType")
                    DocumentSourceReference = dr("DocumentSourceReference")
                    Exit For
                End If
            Next

            Return True
        End If
    End Function

    Private Function ConvertPDFUsingLayoutFiles(ByRef HeaderData As Dictionary(Of String, String), PDFWorkingFile As String, ExcelWorkingFile As String, DocumentType As String, DocumentSourceReference As String) As Boolean
        Dim LayoutCount As Long = 0
        Dim PDFConverted As Boolean = True

        For Each SpecificLayoutFile As String In Directory.GetFiles(Path.Combine(AppDataPath, "Processing - Layouts"), DocumentType & "_" & DocumentSourceReference & "*.layoutx")
            LayoutCount += 1

            Log(LogLevels.Info, "Layout # " + LayoutCount.ToString)
            Log(LogLevels.Info, "Layout File - " + SpecificLayoutFile)

            Dim ExistingFileAction As String
            If LayoutCount = 1 Then
                ExistingFileAction = "replace"
            Else
                ExistingFileAction = "append"
            End If

            Dim PDF2XLArguments As String = "-input=""" + PDFWorkingFile + """ " +
                                            "-layout=""" + SpecificLayoutFile + """ " +
                                            "-format=excelfile " +
                                            "-output=""" + ExcelWorkingFile + """ " +
                                            "-existingfile=" + ExistingFileAction + " " +
                                            "-noui"

            Dim startInfo As New ProcessStartInfo
            startInfo.FileName = My.Settings.PDF2XL_Executable
            startInfo.Arguments = PDF2XLArguments
            startInfo.UseShellExecute = True

            Dim Converter As System.Diagnostics.Process = Process.Start(startInfo)
            Dim timeout As Integer = 60000 '1 minute in milliseconds

            If Not Converter.WaitForExit(timeout) Then
                AddHeaderData(HeaderData, "DocumentValid", "FALSE")
                AddHeaderData(HeaderData, "Exceptions", "Timeout waiting for PDF2XL Conversion.")

                PDFConverted = False
                Exit For
            End If
        Next

        Return PDFConverted
    End Function

    Private Sub ReadContentFromSpreadsheet(ByRef HeaderData As Dictionary(Of String, String), ByRef DetailsData As DataTable, ByRef ExtraTables As Dictionary(Of String, DataTable), ExcelWorkingFile As String)
        Dim wb As Excel.Workbook
        Dim ws As Excel.Worksheet

        wb = ExcelApp.Workbooks.Open(ExcelWorkingFile, [ReadOnly]:=True)
        ' Go through each sheet

        For Each ws In wb.Worksheets
            Dim SheetStartRow As Long = 2
            If My.Settings.PDF2XL_Trial Then SheetStartRow = 5

            Dim CurrentSection As String = ""
            Dim SectionStartRow As Long = 0

            ' Read through sheet and find key field in column 1 (Field/Details/Table_xxx)
            Dim Row As Long

            For Row = SheetStartRow To ws.UsedRange.Rows.Count
                Select Case GetCellValue(ws, Row, 1).ToUpper
                    Case "FIELDS"
                        If SectionStartRow > 0 Then ProcessSection(ws, HeaderData, DetailsData, ExtraTables, CurrentSection, SectionStartRow, Row - 2)
                        SectionStartRow = Row + 1
                        CurrentSection = GetCellValue(ws, Row, 1).ToString
                    Case "DETAILS"
                        If SectionStartRow > 0 Then ProcessSection(ws, HeaderData, DetailsData, ExtraTables, CurrentSection, SectionStartRow, Row - 2)
                        SectionStartRow = Row + 1
                        CurrentSection = GetCellValue(ws, Row, 1).ToString
                    Case Else
                        If Left(GetCellValue(ws, Row, 1).ToUpper, 6) = "TABLE_" Then
                            If SectionStartRow > 0 Then ProcessSection(ws, HeaderData, DetailsData, ExtraTables, CurrentSection, SectionStartRow, Row - 2)
                            SectionStartRow = Row + 1
                            CurrentSection = GetCellValue(ws, Row, 1).ToString
                        End If
                End Select
            Next

            If SectionStartRow > 0 Then ProcessSection(ws, HeaderData, DetailsData, ExtraTables, CurrentSection, SectionStartRow, Row - 1)
        Next
        wb.Close(False)

    End Sub

    Private Sub ProcessSection(ws As Excel.Worksheet, ByRef HeaderData As Dictionary(Of String, String), ByRef DetailsData As DataTable, ByRef ExtraTables As Dictionary(Of String, DataTable), Section As String, StartRow As Long, EndRow As Long)
        Dim dt As New DataTable

        ' Read the contents of each sheet getting the "Fields/Details/Table_XXX" content (Fields -> HeaderData, Details -> DetailsData, Table_xxx -> xxx DataTable)
        Select Case Section.ToUpper
            Case "FIELDS"
                ' If Trial - Ignore PDF2XL Content at Top
                ReadRangeToDataTable(dt, ws, 0, StartRow, EndRow)

                For Row As Long = 0 To dt.Rows.Count - 1
                    AddHeaderData(HeaderData, dt.Rows(Row)(0).ToString, dt.Rows(Row)(1).ToString)
                Next

            Case "DETAILS"
                ReadRangeToDataTable(dt, ws, StartRow, StartRow + 1, EndRow)

                DetailsData = CreateDetailsDataTable()

                For Row As Long = 0 To dt.Rows.Count - 1
                    Dim dr As DataRow = DetailsData.NewRow
                    Dim AddRow As Boolean = True

                    For Column As Integer = 0 To dt.Columns.Count - 1
                        If DetailsData.Columns.Contains(dt.Columns(Column).ColumnName) Then
                            Select Case dt.Columns(Column).ColumnName
                                Case "Qty", "UnitExGST", "UnitIncGST", "ExtendedExGST", "ExtendedIncGST"
                                    Dim Value As String = Replace(dt.Rows(Row)(Column).ToString, "$", "")
                                    If Not IsNumeric(Value) Then
                                        AddRow = False
                                    Else
                                        dr(dt.Columns(Column).ColumnName) = Value
                                    End If
                                Case Else
                                    dr(dt.Columns(Column).ColumnName) = dt.Rows(Row)(Column).ToString
                            End Select
                        End If
                    Next

                    If AddRow Then
                        dr("ID") = Row + 1
                        DetailsData.Rows.Add(dr)
                    End If
                Next

            Case Else
                If Left(Section.ToUpper, 6) = "TABLE_" Then
                    ReadRangeToDataTable(dt, ws, StartRow, StartRow + 1, EndRow)

                    ' Add to Tables Dictionary
                    ExtraTables.Add(Section, dt)
                Else
                    ' TO DO - Unsupported
                    Log(LogLevels.Warning, "The Worksheet/Section '" + Section + "' was not processed as the naming convention is unsupported.")
                End If
        End Select

    End Sub

    Private Sub UpdateDetails(ByRef DetailsData As DataTable)
        ' Validate the Detail Line Contents
        For RowIdx As Long = 0 To DetailsData.Rows.Count - 1
            UpdateMissingDataInDetailRow(DetailsData, RowIdx)
        Next
    End Sub

    Private Sub UpdateMissingDataInDetailRow(ByRef DetailsData As DataTable, RowIdx As Long)
        ' Get Invoice GST Rate (Default 10%) - TO DO - We shouldget this as a setting from the Classification Matrix
        Dim GSTRate As Decimal = 10

        ' Update Missing Data
        With DetailsData.Rows(RowIdx)
            ' **** MISSING EX GST PRICES ***
            ' Missing Unit Ex but have Extended Ex
            If .Item("UnitExGST").ToString = "" And .Item("ExtendedExGST").ToString <> "" Then .Item("UnitExGST") = Math.Round(.Item("ExtendedExGST") / .Item("Qty"), 2)

            ' Missing Extended Ex but have Unit Ex
            If .Item("ExtendedExGST").ToString = "" And .Item("UnitExGST").ToString <> "" Then .Item("ExtendedExGST") = Math.Round(.Item("UnitExGST") * .Item("Qty"), 2)

            ' Need to Assume GST is Included
            ' Missing Unit Ex GST - ExtGST = IncGST / 1.1
            If .Item("UnitExGST").ToString = "" And .Item("UnitIncGST").ToString <> "" Then .Item("UnitExGST") = Math.Round(.Item("UnitIncGST") / ((100 + GSTRate) / 100), 2)

            ' Missing Extended Ex GST - ExtGST = IncGST / 1.1
            If .Item("ExtendedExGST").ToString = "" And .Item("ExtendedIncGST").ToString <> "" Then .Item("ExtendedExGST") = Math.Round(.Item("ExtendedIncGST") / ((100 + GSTRate) / 100), 2)


            ' **** MISSING INC GST PRICES ***
            ' Missing Unit Inc but have Extended Inc
            If .Item("UnitIncGST").ToString = "" And .Item("ExtendedIncGST").ToString <> "" Then .Item("UnitIncGST") = Math.Round(.Item("ExtendedIncGST") / .Item("Qty"), 2)

            ' Missing Extended Inc GST but have Unit Inc GST
            If .Item("ExtendedIncGST").ToString = "" And .Item("UnitIncGST").ToString <> "" Then .Item("ExtendedIncGST") = Math.Round(.Item("UnitIncGST") * .Item("Qty"), 2)

            ' Need to Assume GST is Included
            ' Missing Unit Inc GST but not Unit Ex 
            If .Item("UnitIncGST").ToString = "" And .Item("UnitExGST").ToString <> "" Then .Item("UnitIncGST") = Math.Round(.Item("UnitExGST") * ((100 + GSTRate) / 100), 2)

            ' Missing Extended Inc GST but not Extended Ex
            If .Item("ExtendedIncGST").ToString = "" And .Item("ExtendedExGST").ToString <> "" Then .Item("ExtendedIncGST") = Math.Round(.Item("ExtendedExGST") * ((100 + GSTRate) / 100), 2)

            ' Check if GST Free
            .Item("GSTFree") = (.Item("ExtendedExGST").ToString = .Item("ExtendedIncGST").ToString)
        End With
    End Sub

    Private Sub AddDetailLineIfValid(ByRef DetailsData As DataTable, Code As String, Description As String, Qty As Decimal, UnitExGST As Decimal, UnitIncGST As Decimal, ExtendedExGST As Decimal, ExtendedIncGST As Decimal)

        If IsNumeric(UnitExGST) Or IsNumeric(UnitIncGST) Or IsNumeric(ExtendedExGST) Or IsNumeric(ExtendedIncGST) Then
            If UnitExGST <> 0 Or UnitIncGST <> 0 Or ExtendedExGST <> 0 Or ExtendedIncGST <> 0 Then
                ' Find Last ID
                Dim LastID As Long = Convert.ToInt64(DetailsData.Select("ID=max(ID)")(0)("ID"))

                Dim dr As DataRow = DetailsData.NewRow

                dr("ID") = LastID + 1
                dr("Code") = Code
                dr("Description") = Description
                dr("Qty") = Qty
                dr("UnitExGST") = UnitExGST
                dr("UnitIncGST") = UnitIncGST
                dr("ExtendedExGST") = ExtendedExGST
                dr("ExtendedIncGST") = ExtendedIncGST

                DetailsData.Rows.Add(dr)

                UpdateMissingDataInDetailRow(DetailsData, DetailsData.Rows.Count - 1)
            End If

        End If

    End Sub

    Private Sub UpdateAndValidateHeader(ByRef HeaderData As Dictionary(Of String, String), ByRef DetailsData As DataTable)
        ' Get Invoice GST Rate (Default 10%) - TO DO - We shouldget this as a setting from the Classification Matrix
        Dim GSTRate As Decimal = 10

        ' If Freight or Account Fees In Header than add to Details
        Dim InvFreightIncGST As String = Replace(GetHeaderData(HeaderData, "InvoiceFreightIncGST", ""), "$", "")
        Dim InvFreightExGST As String = Replace(GetHeaderData(HeaderData, "InvoiceFreightExGST", ""), "$", "")
        Dim InvAccFeeIncGST As String = Replace(GetHeaderData(HeaderData, "InvoiceAccountFeeIncGST", ""), "$", "")
        Dim InvAccFeeExGST As String = Replace(GetHeaderData(HeaderData, "InvoiceAccountFeeExGST", ""), "$", "")

        AddDetailLineIfValid(DetailsData, "FREIGHT", "Freight From Invoice Footer", 1, InvFreightExGST, InvFreightIncGST, InvFreightExGST, InvFreightIncGST)
        AddDetailLineIfValid(DetailsData, "ACCOUNTFEE", "Account Keeping Fee From Invoice Footer", 1, InvAccFeeExGST, InvAccFeeIncGST, InvAccFeeExGST, InvAccFeeIncGST)

        ' Check for Required Data (Invoice Totals)
        Dim InvTotIncGST As String = Replace(GetHeaderData(HeaderData, "InvoiceTotalIncGST", ""), "$", "")
        Dim InvTotExGST As String = Replace(GetHeaderData(HeaderData, "InvoiceTotalExGST", ""), "$", "")
        Dim InvTotGST As String = Replace(GetHeaderData(HeaderData, "InvoiceTotalGST", ""), "$", "")

        ' Update/Add the Inoive Totals
        ' Total Inc is Blank and we have Total Ex and Total GST
        If InvTotIncGST = "" And IsNumeric(InvTotExGST) And IsNumeric(InvTotGST) Then
            InvTotIncGST = CDec(InvTotExGST) + CDec(InvTotGST)
            AddHeaderData(HeaderData, "InvoiceTotalIncGST", InvTotIncGST)
        End If

        ' Total Inc & Total GST is Blank and we have Total Ex
        If InvTotIncGST = "" And InvTotGST = "" And IsNumeric(InvTotExGST) Then
            InvTotIncGST = Math.Round(CDec(InvTotExGST) * ((100 + GSTRate) / 100), 2)
            AddHeaderData(HeaderData, "InvoiceTotalIncGST", InvTotIncGST)
        End If

        ' Total Ex is Blank and we have Total Inc and Total GST
        If InvTotExGST = "" And IsNumeric(InvTotIncGST) And IsNumeric(InvTotGST) Then
            InvTotExGST = CDec(InvTotIncGST) - CDec(InvTotGST)
            AddHeaderData(HeaderData, "InvoiceTotalExGST", InvTotIncGST)
        End If

        ' Total Ex & Total GST is Blank and we have Total Inc -> IncGST / 1.1 (10%)
        If InvTotExGST = "" And InvTotGST = "" And IsNumeric(InvTotIncGST) Then
            InvTotExGST = Math.Round(CDec(InvTotIncGST) / ((100 + GSTRate) / 100), 2)
            AddHeaderData(HeaderData, "InvoiceTotalExGST", InvTotIncGST)
        End If

        If InvTotIncGST = "" Or InvTotExGST = "" Then ' If all Totals are empty - invalid invoice
            SetInvoiceInvalid(HeaderData, "Insuffient Invoice Totals Provided - Must Have at Least Total Inc or Excluding GST")
        Else
            ' Get the Total of Invoice from Detail Lines
            Dim SumQty As Decimal
            Dim SumExtendedExGST As Decimal
            Dim SumExtendedIncGST As Decimal

            For RowIdx As Long = 0 To DetailsData.Rows.Count - 1
                With DetailsData.Rows(RowIdx)
                    SumQty += .Item("Qty")
                    SumExtendedExGST += .Item("ExtendedExGST")
                    SumExtendedIncGST += .Item("ExtendedIncGST")
                End With
            Next

            Log(LogLevels.Info, "Sum of Details Qty - " + SumQty.ToString)
            Log(LogLevels.Info, "Sum of Details Extended Ex GST - " + SumExtendedExGST.ToString)
            Log(LogLevels.Info, "Sum of Details Extended Inc GST - " + SumExtendedIncGST.ToString)

            ' Validate Invoice Total Against Details Totals
            If CDec(InvTotIncGST) <> SumExtendedIncGST Then
                SetInvoiceInvalid(HeaderData, "Invoice Total Including GST doesn't match detail lines.")
            ElseIf CDec(InvTotExGST) <> SumExtendedExGST Then
                SetInvoiceInvalid(HeaderData, "Invoice Total Excluding GST doesn't match detail lines.")
            End If

            ' Validate Total Item Count Against Details Qtys (If Available)
            If IsNumeric(GetHeaderData(HeaderData, "InvoiceTotalQty", "")) Then
                If CDec(GetHeaderData(HeaderData, "InvoiceTotalQty", "")) <> SumQty Then
                    SetInvoiceInvalid(HeaderData, "Invoice Total Quantity doesn't match detail lines.")
                End If
            End If
        End If
    End Sub

    Private Sub ReadRangeToDataTable(ByRef dt As DataTable, ws As Excel.Worksheet, HeaderRow As Long, StartRow As Long, EndRow As Long)
        Dim SheetContents As Object

        dt = New DataTable

        If HeaderRow <= 0 Then
            For i As Long = 1 To ws.UsedRange.Columns.Count
                dt.Columns.Add("Field" + i.ToString, GetType(String))
            Next
        Else
            For i As Long = 1 To ws.UsedRange.Columns.Count
                dt.Columns.Add(GetCellValue(ws, HeaderRow, i), GetType(String))
            Next
        End If

        If StartRow <= 0 Then StartRow = 1

        Dim Myrange As Excel.Range = ws.Range(ws.Cells(StartRow, 1), ws.Cells(EndRow, ws.UsedRange.Columns.Count))

        Myrange.UnMerge()

        ' Get the Sheet Contents into Array for fast reading
        SheetContents = Myrange.Value

        ' Go through each row 
        For Row As Long = 1 To UBound(SheetContents, 1)
            Dim dr As DataRow = dt.NewRow

            ' With each Column 
            For Column As Long = 1 To UBound(SheetContents, 2)
                Dim ColIdx As Integer = Column - 1
                If IsNothing(SheetContents(Row, Column)) Then
                    dr(ColIdx) = ""
                Else
                    dr(ColIdx) = SheetContents(Row, Column).ToString
                End If
            Next
            dt.Rows.Add(dr)
        Next

    End Sub

    Private Sub SetInvoiceInvalid(d As Dictionary(Of String, String), Exception As String)
        If d.ContainsKey("DocumentValid") Then
            d("DocumentValid") = "FALSE"
        Else
            d.Add("DocumentValid", "FALSE")
        End If

        If d.ContainsKey("Exceptions") Then
            d("Exceptions") = d("Exceptions") + Environment.NewLine + Exception
        Else
            d.Add("Exceptions", Exception)
        End If

    End Sub

    Private Function GetDocumentClassificationMatrix(Spreadsheet As String) As Data.DataTable
        Dim wb As Excel.Workbook
        Dim ws As Excel.Worksheet

        wb = ExcelApp.Workbooks.Open(Spreadsheet, [ReadOnly]:=True)

        ' Set the Worksheet
        ws = wb.Sheets(1)

        Dim dt As New Data.DataTable

        ' Create typed columns in the DataTable.
        dt.Columns.Add("Priority", GetType(String))
        dt.Columns.Add("DocumentType", GetType(String))
        dt.Columns.Add("DocumentSourceReference", GetType(String))
        dt.Columns.Add("Description", GetType(String))

        dt.Columns.Add("SearchText1", GetType(String))
        dt.Columns.Add("SearchText2", GetType(String))
        dt.Columns.Add("SearchText3", GetType(String))

        ' Get the Sheet Contents into Array for fast reading
        Dim SheetContents As Object
        SheetContents = ws.Range(ws.Cells(2, 1), ws.Cells(ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count)).Value

        ' Read Rows
        For Row As Long = 1 To UBound(SheetContents)
            Dim dr As DataRow = dt.NewRow

            dr("Priority") = SheetContents(Row, 1)
            dr("DocumentType") = SheetContents(Row, 2)
            dr("DocumentSourceReference") = SheetContents(Row, 3)
            dr("Description") = SheetContents(Row, 4)

            dr("SearchText1") = SheetContents(Row, 5)
            dr("SearchText2") = SheetContents(Row, 6)
            dr("SearchText3") = SheetContents(Row, 7)

            dt.Rows.Add(dr)
        Next

        wb.Close(False)

        Return dt
    End Function

    Private Function CreateDetailsDataTable() As Data.DataTable
        Dim dt As New Data.DataTable

        ' Create typed columns in the DataTable.
        dt.Columns.Add("ID", GetType(Integer))
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Description", GetType(String))
        dt.Columns.Add("Qty", GetType(Decimal))
        dt.Columns.Add("UnitExGST", GetType(Decimal))
        dt.Columns.Add("UnitIncGST", GetType(Decimal))
        dt.Columns.Add("ExtendedExGST", GetType(Decimal))
        dt.Columns.Add("ExtendedIncGST", GetType(Decimal))
        dt.Columns.Add("GSTFree", GetType(Boolean))

        Return dt
    End Function

    Private Function GetPostConversionTasks(Spreadsheet As String) As Data.DataTable
        Dim wb As Excel.Workbook
        Dim ws As Excel.Worksheet

        Dim dt As New Data.DataTable

        ' Create typed columns in the DataTable.
        dt.Columns.Add("Action", GetType(String))
        dt.Columns.Add("Source", GetType(String))
        dt.Columns.Add("FieldName", GetType(String))
        dt.Columns.Add("Comment", GetType(String))

        dt.Columns.Add("ContentSearchExp", GetType(String))
        dt.Columns.Add("ContentColumn", GetType(String))
        dt.Columns.Add("ContentFindText", GetType(String))
        dt.Columns.Add("ContentReplaceText", GetType(String))
        dt.Columns.Add("ContentMaxChars", GetType(String))

        If File.Exists(Spreadsheet) Then
            wb = ExcelApp.Workbooks.Open(Spreadsheet, [ReadOnly]:=True)

            ' Set the Worksheet
            ws = wb.Sheets(1)

            ' Get the Sheet Contents into Array for fast reading
            Dim SheetContents As Object
            SheetContents = ws.Range(ws.Cells(3, 1), ws.Cells(ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count)).Value

            ' Read Rows
            For Row As Long = 1 To UBound(SheetContents)
                Dim dr As DataRow = dt.NewRow

                dr("Action") = SheetContents(Row, 1)
                dr("Source") = SheetContents(Row, 2)
                dr("FieldName") = SheetContents(Row, 3)
                dr("Comment") = SheetContents(Row, 4)

                dr("ContentSearchExp") = SheetContents(Row, 5)
                dr("ContentColumn") = SheetContents(Row, 6)
                dr("ContentFindText") = SheetContents(Row, 7)
                dr("ContentReplaceText") = SheetContents(Row, 8)
                dr("ContentMaxChars") = SheetContents(Row, 9)

                dt.Rows.Add(dr)
            Next

            wb.Close(False)
        End If

        Return dt
    End Function


    Private Sub CleanInvoiceDetails(ByRef DetailsData As DataTable, dr As DataRow)
        Dim SearchExpression As String = dr("ContentSearchExp")
        SearchExpression = Replace(SearchExpression, """", "'")

        Log(LogLevels.Info, "Cleaning Invoice Details (" + SearchExpression + ")")

        DetailsData = DetailsData.Select(SearchExpression).CopyToDataTable()

        Log(LogLevels.Info, "Invoice Details Remaining - " + DetailsData.Rows.Count.ToString)
    End Sub

    Private Sub UpdateHeaderContent(ByRef HeaderData As Dictionary(Of String, String), DetailsData As DataTable, ExtraTables As Dictionary(Of String, DataTable), dr As DataRow)
        Dim dt As New DataTable

        Select Case dr("Action").ToString.ToUpper
            Case "Get Header Content".ToUpper
                Dim SearchExpression As String = dr("ContentSearchExp").ToString
                SearchExpression = Replace(SearchExpression, """", "'")
                Dim Source As String = dr("Source").ToString

                If Source.ToUpper = "DETAILS" Then
                    dt = DetailsData
                ElseIf Left(Source.ToUpper, 6) = "TABLE_" Then
                    If ExtraTables.ContainsKey(Source) Then
                        dt = ExtraTables(Source)
                    Else
                        Log(LogLevels.Warning, "Unable to find Source Content '" + Source + "'")
                    End If
                End If

                Log(LogLevels.Info, "Getting Header Content (" + SearchExpression + ")")

                Try
                    dt = dt.Select(SearchExpression).CopyToDataTable()

                    If dt.Rows.Count > 0 Then
                        Log(LogLevels.Info, "Getting Header Content - Rows = " + dt.Rows.Count.ToString)

                        ' Get the Field Number/Name and ther Max Chars
                        Dim Field As String = dr("ContentColumn").ToString
                        Dim MaxChars As String = dr("ContentMaxChars").ToString
                        Dim Value As String = ""

                        If IsNumeric(Field) Then
                            Dim Column As Integer = CInt(Field) - 1
                            Value = dt.Rows(0)(Column).ToString
                        Else
                            Value = dt.Rows(0)(Field).ToString
                        End If

                        AddHeaderData(HeaderData, dr("FieldName").ToString, Value)
                    Else
                        Log(LogLevels.Warning, "0 rows returned - Header not updated.")
                    End If

                Catch ex As Exception
                    Log(LogLevels.Warning, "Error retrieving content from Source Table - Error: " + ex.Message)

                End Try

            Case "Replace Header Content".ToUpper
                Dim FieldName As String = dr("FieldName").ToString
                Dim FindText As String = dr("ContentFindText").ToString
                Dim Replacetext As String = dr("ContentReplaceText").ToString

                ' Get the Existing Value
                Dim value As String = GetHeaderData(HeaderData, FieldName)
                ' Replace the text
                value = Replace(value, FindText, Replacetext)
                ' Update the Header Data
                AddHeaderData(HeaderData, FieldName, value)

        End Select

    End Sub


    Private Function GetCompanyName(ABN As String) As String
        Dim Search As CompanySearch_httpXMLSearch
        Dim SearchPayload As String
        Dim MainName As String = ""
        Dim MainTradingName As String = ""
        Dim CompanyName As String = ""

        Log(LogLevels.Trace, "GetCompanyName - Declaring Search Variable...")
        Search = New CompanySearch_httpXMLDocumentSearch

        Log(LogLevels.Trace, "GetCompanyName - Performing ABN Search...")
        SearchPayload = Search.ABNSearch(ABN, "n", "371f387c-0a3b-420b-ac39-00bb04f5b85f")

        Log(LogLevels.Trace, "GetCompanyName - Parsing Result...")
        Dim p As New XMLMessageParser(SearchPayload)

        Return p.GetCompanyName()

    End Function


End Module
