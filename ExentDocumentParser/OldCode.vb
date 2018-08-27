Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Text.RegularExpressions

Module OldCode
    'Enum FieldTypes
    '    [String] = 0
    '    Currency = 1
    '    [Date] = 2
    '    Quantity = 3
    '    ABN = 4
    'End Enum

    'Dim ExcelApp As Excel.Application
    'Dim f As FormActivity
    'Dim ShowOfficeApplications As Boolean

    'Dim AppDataPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Exent\DocumentParser\")
    'Dim WorkingPath As String = Path.Combine(AppDataPath, "Working")

    'Public Sub GetDocumentContents(PDFDocument As String, ExcelWorkbook As String, WordDocument As String, ByRef HeaderData As Dictionary(Of String, String), ByRef DetailsData As DataTable, Optional DisplayOfficeApps As Boolean = False)
    '    Dim wb As Excel.Workbook
    '    Dim ws As Excel.Worksheet
    '    Dim PDFWorkingFile As String = ""
    '    'Dim ProcessingData As Data.DataTable

    '    f = New FormActivity
    '    f.Show()

    '    ' ***************************************************************************
    '    ' Make sure App Data Working Path exists
    '    If Not Directory.Exists(WorkingPath) Then Directory.CreateDirectory(WorkingPath)
    '    'Make sure working path is empty
    '    For Each file As String In Directory.GetFiles(WorkingPath)
    '        IO.File.Delete(file)
    '    Next

    '    ' ***************************************************************************
    '    ' Load Excel Application
    '    UpdateStatus(f, "Loading Excel...")
    '    ExcelApp = New Excel.Application
    '    ExcelApp.Visible = DisplayOfficeApps

    '    ' ***************************************************************************
    '    '
    '    If PDFDocument.Length > 0 Then
    '        UpdateStatus(f, "Converting PDF...")
    '        PDFWorkingFile = Path.Combine(WorkingPath, Path.GetFileName(PDFDocument))
    '        ' Move file to Working Folder
    '        File.Copy(PDFDocument, PDFWorkingFile, True)
    '        PDFDocument = PDFWorkingFile

    '    ElseIf ExcelWorkbook.Length > 0 Then
    '        UpdateStatus(f, "Converting Excel Workbook...")

    '        MsgBox("Excel Conversion Not Currently Supported - Please contact Exent Support.")

    '        PDFWorkingFile = Path.Combine(WorkingPath, Path.GetFileNameWithoutExtension(ExcelWorkbook), ".pdf")

    '        ' TO DO - Open using Excel and Export to PDF

    '    ElseIf WordDocument.Length > 0 Then
    '        UpdateStatus(f, "Converting Word Document...")

    '        MsgBox("Word Conversion Not Currently Supported - Please contact Exent Support.")

    '        PDFWorkingFile = Path.Combine(WorkingPath, Path.GetFileNameWithoutExtension(WordDocument), ".pdf")

    '        ' TO DO - Open using Word and Export to PDF
    '    End If

    '    ' ***************************************************************************
    '    ' If PDF Document Exists - Extract Information
    '    If PDFWorkingFile.Length > 0 Then

    '        ' **** Create a new Temp Workbook in Excel and Add PDF File as OLE Object to find the Page Size of the PDF
    '        UpdateStatus(f, "Getting PDF Page Orientation...")
    '        wb = ExcelApp.Workbooks.Add()
    '        ws = wb.Worksheets(1)

    '        ws.OLEObjects.Add(Filename:=PDFWorkingFile, Link:=False, DisplayAsIcon:=False).Select

    '        Dim PDFHeight As Long = ws.OLEObjects(1).Height
    '        Dim PDFWidth As Long = ws.OLEObjects(1).Width
    '        Dim IsPDFPortrait As Boolean = PDFHeight > PDFWidth

    '        wb.Close(False)


    '        ' *** Classify PDF Document - Get the Document Type and Source Information
    '        UpdateStatus(f, "Classifying PDF Document...")
    '        ' Convert contents of PDF Invoice using PDF2XL with Full Document Layout for Orientation to Excel
    '        Dim LayoutFile As String
    '        If IsPDFPortrait Then
    '            LayoutFile = Path.Combine(AppDataPath, "Processing - Layouts", "AllContent_Portrait.layoutx")
    '        Else
    '            LayoutFile = Path.Combine(AppDataPath, "Processing - Layouts", "AllContent_Landscape.layoutx")
    '        End If

    '        Dim ExcelWorkingFile As String = Path.Combine(WorkingPath, "FileContents.xlsx")

    '        Dim PDF2XLArguments As String = "-input=""" + PDFWorkingFile + """ " +
    '                                        "-layout=""" + LayoutFile + """ " +
    '                                        "-format=excelfile " +
    '                                        "-output=""" + ExcelWorkingFile + """ " +
    '                                        "-existingfile=replace " +
    '                                        "-noui"

    '        Dim startInfo As New ProcessStartInfo
    '        startInfo.FileName = "C:\Program Files (x86)\CogniView\PDF2XL\PDF2XL.exe"
    '        startInfo.Arguments = PDF2XLArguments
    '        startInfo.UseShellExecute = True
    '        Dim Converter As System.Diagnostics.Process = Process.Start(startInfo)
    '        Dim timeout As Integer = 60000 '1 minute in milliseconds

    '        If Not Converter.WaitForExit(timeout) Then
    '            ' TO DO - Report Issue
    '        Else
    '            ' Get the Document Classification Matrix & Search for content
    '            Dim DocumentType As String = "Invoice"
    '            Dim DocumentSourceReference As String = "TheBossShop"

    '            ' TO DO - Open Document Classification Matrix Spreadsheet and read into DataTable
    '            ' TO DO - Open Excel Working File Spreadsheet
    '            ' TO DO - Iterate through DataTable looking for match
    '            ' TO DO - Close Excel Files

    '            Log(LogLevels.Info, "Document Type - " + DocumentType)
    '            Log(LogLevels.Info, "Document Source Reference - " + DocumentSourceReference)

    '            If DocumentType.Length = 0 Or DocumentSourceReference.Length = 0 Then
    '                ' TO DO - Report Issue
    '            Else
    '                ' *** Read PDF Document
    '                UpdateStatus(f, "Reading Contents of PDF using Document/Source Layout(s)...")
    '                ' Convert contents of PDF Invoice using PDF2XL with the Specified Layout(s) for Document Type & Source to Excel
    '                Dim LayoutCount As Long = 0
    '                Dim PDFConverted As Boolean = True

    '                For Each SpecificLayoutFile As String In Directory.GetFiles(Path.Combine(AppDataPath, "Processing - Layouts"), DocumentType & "_" & DocumentSourceReference & "*.layoutx")
    '                    LayoutCount += 1

    '                    Log(LogLevels.Info, "Layout # " + LayoutCount.ToString)
    '                    Log(LogLevels.Info, "Layout File - " + SpecificLayoutFile)

    '                    Dim ExistingFileAction As String
    '                    If LayoutCount = 1 Then
    '                        ExistingFileAction = "replace"
    '                    Else
    '                        ExistingFileAction = "append"
    '                    End If

    '                    PDF2XLArguments = "-input=""" + PDFWorkingFile + """ " +
    '                            "-layout=""" + SpecificLayoutFile + """ " +
    '                            "-format=excelfile " +
    '                            "-output=""" + ExcelWorkingFile + """ " +
    '                            "-existingfile=" + ExistingFileAction + " " +
    '                            "-noui"

    '                    startInfo.FileName = "C:\Program Files (x86)\CogniView\PDF2XL\PDF2XL.exe"
    '                    startInfo.Arguments = PDF2XLArguments
    '                    startInfo.UseShellExecute = True

    '                    Converter = Process.Start(startInfo)

    '                    If Not Converter.WaitForExit(timeout) Then
    '                        PDFConverted = False
    '                        Exit For
    '                    End If
    '                Next

    '                If PDFConverted Then
    '                    ' *** Read data into DataTables from Spreadsheet
    '                    ' Open Excel Working File Spreadsheet
    '                    wb = ExcelApp.Workbooks.Open(ExcelWorkingFile, [ReadOnly]:=True)
    '                    ' Go through each sheet
    '                    Dim dt As New DataTable
    '                    Dim StartRow As Long = 0

    '                    For Each ws In wb.Worksheets
    '                        Select Case ws.Name.ToUpper
    '                            Case "FIELDS"
    '                                ReadRangeToDataTable(dt, ExcelWorkingFile, "Fields")

    '                                For r As Long = StartRow To dt.Rows.Count - 1


    '                                Next
    '                                MsgBox(dt.Rows.Count)

    '                            Case "DETAILS"
    '                                ReadExcelFile(dt, ExcelWorkingFile, "Details")
    '                                MsgBox(dt.Rows.Count)

    '                            Case Else
    '                                If Left(ws.Name.ToUpper, 6) = "TABLE_" Then

    '                                Else
    '                                    ' TO DO - Unsupported

    '                                End If
    '                        End Select
    '                    Next
    '                    wb.Close()

    '                    ' TO DO - Read the contents of each sheet getting the "Fields/Details/Table_XXX" content (Fields -> HeaderData, Details -> DetailsData, Table_xxx -> xxx DataTable)

    '                    ' *** Run Post Conversion Tasks
    '                    ' Open the tasks spreadsheet for the Document Type & Source 

    '                    ' Apply each of the tasks to the DataTables



    '                End If
    '            End If
    '        End If
    '    End If

    '    UpdateStatus(f, "Closing Excel...")
    '    ExcelApp.Quit()

    '    releaseObject(ExcelApp)
    'End Sub

    'Private Sub ReadRangeToDataTable(ByRef dt As DataTable, ws As Excel.Worksheet, HeaderRow As Long, StartRow As Long)
    '    Dim SheetContents As Object

    '    dt = New DataTable

    '    If HeaderRow <= 0 Then
    '        For i As Long = 1 To ws.UsedRange.Columns.Count
    '            dt.Columns.Add("Field" + i)
    '        Next
    '    Else
    '        For i As Long = 1 To ws.UsedRange.Columns.Count
    '            dt.Columns.Add(ws.Cells(HeaderRow, i).value)
    '        Next
    '    End If

    '    If StartRow <= 0 Then StartRow = 1

    '    Dim Myrange As Excel.Range = ws.Range(ws.Cells(StartRow, 1), ws.Cells(ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count))

    '    Myrange.UnMerge()

    '    ' Get the Sheet Contents into Array for fast reading
    '    SheetContents = Myrange.Value



    'End Sub

    'Private Sub OldFunction()
    '    '' See if the Excel Workbook already exists (all converted - most likely only going to ne used in testing)
    '    'If File.Exists(ExcelWorkbook) Then
    '    '    UpdateStatus(f, "Loading Invoice Spreadsheet...")
    '    '    wb = ExcelApp.Workbooks.Open(ExcelWorkbook)
    '    'ElseIf WordDocument <> "" Then
    '    '    UpdateStatus(f, "Converting Content to Spreadsheet...")

    '    '    ' Create New Excel Document 
    '    '    wb = ExcelApp.Workbooks.Add
    '    '    wb.SaveAs(ExcelWorkbook)

    '    '    ConvertDOCXToXLSX(f, WordDocument, ExcelApp, wb, DisplayOfficeApps)
    '    'Else
    '    '    ' Nothing to Do...

    '    '    Exit Sub
    '    'End If

    '    '' Set the Worksheet
    '    'ws = wb.ActiveSheet

    '    '' Remove Merged Cells
    '    ''UnMergeCells(ws)

    '    '' ***************************************************************************
    '    'Dim Myrange As Excel.Range = ws.Range(ws.Cells(1, 1), ws.Cells(ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count))

    '    'Myrange.UnMerge()

    '    '' Get the Sheet Contents into Array for fast reading
    '    'SheetContents = Myrange.Value

    '    '' Load the Data from the Default Template
    '    'UpdateStatus(f, "Loading the Default Processing Template...")
    '    'ProcessingTemplate = ProcessingTemplatesFolder + "Default.xlsx"
    '    'ProcessingData = GetProcessingTemplateData(ProcessingTemplate)

    '    'Log(LogLevels.Trace, "Setting Defaults for Invoice...")
    '    'AddHeaderData(HeaderData, "InvoiceValid", "TRUE")
    '    'AddHeaderData(HeaderData, "Exceptions", "")

    '    '' ***************************************************************************
    '    '' **** Get ABN **** REQUIRED ****
    '    'Log(LogLevels.Trace, "Getting ABN...")
    '    'ABN = GetValueFromSpreadsheet("ABN", FieldTypes.ABN, ProcessingData)
    '    '' Remnove Spaces from ABN
    '    'ABN = ABN.Replace(" ", "")
    '    'AddHeaderData(HeaderData, "ABN", ABN)
    '    'If ABN = "" Then SetInvoiceInvalid(HeaderData, "Missing ABN!")
    '    'Log(LogLevels.Trace, "ABN = " + ABN)

    '    '' **** Get Company Name ****
    '    '' Use Web Service
    '    'If ABN <> "" Then
    '    '    Log(LogLevels.Trace, "Getting Company Name...")
    '    '    Value = GetCompanyName(ABN)
    '    '    Log(LogLevels.Trace, "Company Name = " + Value)
    '    'Else
    '    '    Value = ""
    '    'End If
    '    'AddHeaderData(HeaderData, "InvoiceCompany", Value)

    '    '' ***************************************************************************
    '    '' **** Check for Custom Processing Template ****
    '    'Log(LogLevels.Trace, "Checking for Processing Templates Folder...")
    '    'If Directory.GetFiles(ProcessingTemplatesFolder, ABN + "*.xlsx").Count > 0 Then
    '    '    UpdateStatus(f, "Loading the Custom Processing Template...")
    '    '    ' Get the first file in the list (should only be 1)
    '    '    ProcessingTemplate = Directory.GetFiles(ProcessingTemplatesFolder, ABN + "*.xlsx")(0)
    '    '    ' Load the Custom Processing Template
    '    '    ProcessingData = GetProcessingTemplateData(ProcessingTemplate)
    '    'End If

    '    '' ***************************************************************************
    '    '' **** Clean the Detail Lines ****
    '    'Dim DetailLinesCleaned As Boolean = False

    '    'UpdateStatus(f, "Cleaning Invoice Detail Lines...")
    '    'DetailLinesCleaned = ProcessInvoiceDetails(True, ws, HeaderData, DetailsData, ProcessingData)

    '    'If DetailLinesCleaned Then
    '    '    ' Load the Sheet Contents Again
    '    '    Myrange = ws.Range(ws.Cells(1, 1), ws.Cells(ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count))

    '    '    ' Get the Sheet Contents into Array for fast reading
    '    '    SheetContents = Myrange.Value
    '    'End If

    '    'UpdateStatus(f, "Closing Excel...")
    '    'wb.Save()
    '    'wb.Close()
    '    'ExcelApp.Quit()

    '    'releaseObject(ExcelApp)
    '    'releaseObject(wb)

    '    '' ***************************************************************************
    '    'UpdateStatus(f, "Reading Invoice...")
    '    '' **** Get Invoice Number **** REQUIRED ****
    '    'Log(LogLevels.Trace, "Getting Invoice Number...")
    '    'Value = GetValueFromSpreadsheet("InvoiceNumber", FieldTypes.String, ProcessingData)
    '    'Log(LogLevels.Trace, "Invoice Number = " + Value)
    '    'AddHeaderData(HeaderData, "InvoiceNumber", Value)
    '    'If Value = "" Then SetInvoiceInvalid(HeaderData, "Missing Invoice Number!")

    '    '' **** Get Order Number **** REQUIRED ****
    '    'Value = GetValueFromSpreadsheet("OrderNumber", FieldTypes.String, ProcessingData)
    '    'AddHeaderData(HeaderData, "OrderNumber", Value)
    '    'If Value = "" Then SetInvoiceInvalid(HeaderData, "Missing Order Number!")

    '    '' **** Get Date ****
    '    'Value = GetValueFromSpreadsheet("InvoiceDate", FieldTypes.Date, ProcessingData)
    '    '' Convert to Standard Format
    '    'If IsDate(Value) Then
    '    '    Value = CDate(Value).ToString("d/MM/yyyy")
    '    'Else
    '    '    Value = ""
    '    'End If
    '    'AddHeaderData(HeaderData, "InvoiceDate", Value)

    '    '' **** Get Due Date ****
    '    'Value = GetValueFromSpreadsheet("InvoiceDueDate", FieldTypes.Date, ProcessingData)
    '    '' Convert to Standard Format
    '    'If IsDate(Value) Then
    '    '    Value = CDate(Value).ToString("d/MM/yyyy")
    '    'Else
    '    '    Value = ""
    '    'End If
    '    'AddHeaderData(HeaderData, "InvoiceDueDate", Value)

    '    '' **** Get Terms ****
    '    'Value = GetValueFromSpreadsheet("InvoiceTerms", FieldTypes.String, ProcessingData)
    '    '' Convert to Standard Format
    '    'If Value <> "" Then
    '    '    ' Get numerical numbers only (check if between 1 and 100 - 7/14/30/60/90)
    '    '    Dim TermsDays As String = Regex.Replace(Value, "[^0-9 ]", "")
    '    '    ' Get the first 3 characters only - in case other numbers in string
    '    '    TermsDays = Strings.Left(TermsDays.Trim, 3).Trim
    '    '    ' Get the Text Only to find if Days/Invoice/Statement 
    '    '    Dim TermsPeriod As String = Regex.Replace(Value, "[^a-zA-Z ]", "")

    '    '    If TermsPeriod.ToUpper.Contains("DAY") Or IsNumeric(TermsDays) Then Value = TermsDays + "DAYS"
    '    '    If TermsPeriod.ToUpper.Contains("INV") Or TermsPeriod.ToUpper.Contains("NET") Then Value = Value + " INVOICE"
    '    '    If TermsPeriod.ToUpper.Contains("MONTH") Or TermsPeriod.ToUpper.Contains("EOM") Then Value = Value + " EOM"
    '    '    If TermsPeriod.ToUpper.Contains("ST") Then Value = Value + " STATEMENT"
    '    'End If
    '    'AddHeaderData(HeaderData, "InvoiceTerms", Value)

    '    '' **** Get Total Value of Invoice **** REQUIRED ****
    '    'Value = GetValueFromSpreadsheet("InvoiceTotal", FieldTypes.Currency, ProcessingData)
    '    'AddHeaderData(HeaderData, "InvoiceTotal", Value)
    '    'If IsNumeric(Value) Then
    '    '    InvoiceTotal = CDec(Value)
    '    'Else
    '    '    SetInvoiceInvalid(HeaderData, "Missing Invoice Total!")
    '    'End If

    '    '' **** Get Total Value of GST ****
    '    'Value = GetValueFromSpreadsheet("InvoiceGST", FieldTypes.Currency, ProcessingData)
    '    'AddHeaderData(HeaderData, "InvoiceGST", Value)

    '    'ProcessInvoiceDetails(False, Nothing, HeaderData, DetailsData, ProcessingData)

    '    '' Validate the Detail Line Contents
    '    'For RowIdx As Long = 0 To DetailsData.Rows.Count - 1
    '    '    If DetailsData.Rows(RowIdx)("Qty").ToString = "" Then
    '    '        SetInvoiceInvalid(HeaderData, "Missing Invoice Details - No Quantiy for Line " + (RowIdx + 1).ToString + "!")
    '    '    ElseIf DetailsData.Rows(RowIdx)("UnitExGST").ToString = "" And DetailsData.Rows(RowIdx)("UnitIncGST").ToString = "" And DetailsData.Rows(RowIdx)("ExtendedExGST").ToString = "" And DetailsData.Rows(RowIdx)("ExtendedIncGST").ToString = "" Then
    '    '        SetInvoiceInvalid(HeaderData, "Missing Invoice Details - No Pricing for Line " + (RowIdx + 1).ToString + "!")
    '    '    Else

    '    '        ' Get Invoice GST Rate (Default 10%) - TO DO - Should we validate this against the Doc Total and GDT Total
    '    '        Dim GSTRate As Decimal = 10

    '    '        ' Update Missing Data
    '    '        ' Missing Unit Ex GST & Inc GST but not ExtendedExGST
    '    '        If DetailsData.Rows(RowIdx)("UnitExGST").ToString = "" And DetailsData.Rows(RowIdx)("UnitIncGST").ToString = "" And DetailsData.Rows(RowIdx)("ExtendedExGST").ToString <> "" Then
    '    '            DetailsData.Rows(RowIdx)("UnitExGST") = Math.Round(DetailsData.Rows(RowIdx)("ExtendedExGST") / DetailsData.Rows(RowIdx)("Qty"), 2)
    '    '        End If

    '    '        ' Missing Unit Ex GST & Inc GST but not ExtendedIncGST
    '    '        If DetailsData.Rows(RowIdx)("UnitExGST").ToString = "" And DetailsData.Rows(RowIdx)("UnitIncGST").ToString = "" And DetailsData.Rows(RowIdx)("ExtendedIncGST").ToString <> "" Then
    '    '            DetailsData.Rows(RowIdx)("UnitIncGST") = Math.Round(DetailsData.Rows(RowIdx)("ExtendedIncGST") / DetailsData.Rows(RowIdx)("Qty"), 2)
    '    '        End If

    '    '        ' Missing Unit Ex GST & Inc GST but not ExtendedExGST
    '    '        If DetailsData.Rows(RowIdx)("ExtendedExGST").ToString = "" And DetailsData.Rows(RowIdx)("ExtendedIncGST").ToString = "" And DetailsData.Rows(RowIdx)("UnitExGST").ToString <> "" Then
    '    '            DetailsData.Rows(RowIdx)("ExtendedExGST") = Math.Round(DetailsData.Rows(RowIdx)("UnitExGST") * DetailsData.Rows(RowIdx)("Qty"), 2)
    '    '        End If

    '    '        ' Missing Extended Ex GST & Exetended Inc GST but not Unit Inc GST
    '    '        If DetailsData.Rows(RowIdx)("ExtendedExGST").ToString = "" And DetailsData.Rows(RowIdx)("ExtendedIncGST").ToString = "" And DetailsData.Rows(RowIdx)("UnitIncGST").ToString <> "" Then
    '    '            DetailsData.Rows(RowIdx)("ExtendedIncGST") = Math.Round(DetailsData.Rows(RowIdx)("UnitIncGST") * DetailsData.Rows(RowIdx)("Qty"), 2)
    '    '        End If

    '    '        ' Missing Unit Ex GST - ExtGST = IncGST / 1.1
    '    '        If DetailsData.Rows(RowIdx)("UnitExGST").ToString = "" And DetailsData.Rows(RowIdx)("UnitIncGST").ToString <> "" Then
    '    '            DetailsData.Rows(RowIdx)("UnitExGST") = Math.Round(DetailsData.Rows(RowIdx)("UnitIncGST") / ((100 + GSTRate) / 100), 2)

    '    '            ' Missing Unit Inc GST - IncGST = ExtGST * 1.1
    '    '        ElseIf DetailsData.Rows(RowIdx)("UnitExGST").ToString <> "" And DetailsData.Rows(RowIdx)("UnitIncGST").ToString = "" Then
    '    '            DetailsData.Rows(RowIdx)("UnitIncGST") = Math.Round(DetailsData.Rows(RowIdx)("UnitExGST") * ((100 + GSTRate) / 100), 2)
    '    '        End If

    '    '        ' Missing Exetended Ex GST - ExtGST = IncGST / 1.1
    '    '        If DetailsData.Rows(RowIdx)("ExtendedExGST").ToString = "" And DetailsData.Rows(RowIdx)("ExtendedIncGST").ToString <> "" Then
    '    '            DetailsData.Rows(RowIdx)("ExtendedExGST") = Math.Round(DetailsData.Rows(RowIdx)("ExtendedIncGST") / ((100 + GSTRate) / 100), 2)

    '    '            ' Missing Unit Inc GST - IncGST = ExtGST * 1.1
    '    '        ElseIf DetailsData.Rows(RowIdx)("ExtendedExGST").ToString <> "" And DetailsData.Rows(RowIdx)("ExtendedIncGST").ToString = "" Then
    '    '            DetailsData.Rows(RowIdx)("ExtendedIncGST") = Math.Round(DetailsData.Rows(RowIdx)("ExtendedExGST") * ((100 + GSTRate) / 100), 2)
    '    '        End If
    '    '    End If
    '    'Next

    '    'f.Close()

    'End Sub

    'Private Function ProcessInvoiceDetails(CleaningLinesInSpreadsheet As Boolean, ws As Excel.Worksheet, ByRef HeaderData As Dictionary(Of String, String), ByRef DetailsData As DataTable, ProcessingData As Data.DataTable) As Boolean
    '    ' **** Get Invoice Details ****
    '    Dim HeaderRow As Long = 0       ' Row Number that Header Starts
    '    Dim HeaderRowSpan As Long = 1   ' How many rows in Header
    '    Dim DetailsRowSpan As Long = 1         ' How many rows in Detail Lines
    '    Dim LastRow As Long = 0         ' Last Row of Details
    '    Dim Processed As Boolean = False

    '    DetailsData = CreasteDetailsDataTable()

    '    ' Get the Start Row of Details
    '    HeaderRow = GetRowFromSpreadsheet("InvoiceHeader", FieldTypes.String, ProcessingData, HeaderRowSpan)

    '    If HeaderRow > 0 Then

    '        ' Get the End Row of Details
    '        LastRow = GetRowFromSpreadsheet("InvoiceDetailsEnd", FieldTypes.String, ProcessingData) - 1

    '        ' Get the Row Span of Details
    '        GetRowFromSpreadsheet("InvoiceDetails", FieldTypes.String, ProcessingData, DetailsRowSpan)

    '        ' Get the number of cells populated on first line, we expect all lines below to be the same or within 80%
    '        Dim FirstLineCellsPopulated As Long = 0
    '        If CleaningLinesInSpreadsheet Then
    '            For DetailColumn = 1 To UBound(SheetContents, 2)
    '                ' Cell Has Content
    '                If Not IsNothing(SheetContents(HeaderRow + HeaderRowSpan, DetailColumn)) Then
    '                    ' If content is not blank 
    '                    If SheetContents(HeaderRow + HeaderRowSpan, DetailColumn).ToString.Trim <> "" Then
    '                        FirstLineCellsPopulated = FirstLineCellsPopulated + 1
    '                    End If
    '                End If
    '            Next
    '        End If

    '        If LastRow > 0 Then
    '            If CleaningLinesInSpreadsheet Then
    '                ' Clean up Details Rows (Merge Descriptions Across 2 Lines & Remove Columns with no data)
    '                Processed = CleanLinesInSpreadsheet(ws, HeaderRow, HeaderRowSpan, DetailsRowSpan, LastRow, FirstLineCellsPopulated)
    '            Else
    '                Processed = GetDataFromDetails(HeaderData, DetailsData, ProcessingData, HeaderRow, HeaderRowSpan, DetailsRowSpan, LastRow)
    '            End If
    '        Else
    '            SetInvoiceInvalid(HeaderData, "Missing Invoice Details - Can Not find end of Details!")
    '        End If
    '    Else
    '        SetInvoiceInvalid(HeaderData, "Missing Invoice Details - Can Not find Header!")
    '    End If

    '    Return Processed
    'End Function

    'Private Function CleanLinesInSpreadsheet(ws As Excel.Worksheet, HeaderRow As Long, HeaderRowSpan As Long, DetailsRowSpan As Long, LastRow As Long, FirstLineCellsPopulated As Long) As Boolean
    '    Dim CleanedLines As Boolean = False

    '    ExcelApp.ScreenUpdating = False

    '    ' Merge Descriptions Across 2 Lines  
    '    For DetailRow As Long = HeaderRow + HeaderRowSpan To LastRow Step DetailsRowSpan
    '        ' Get the number of cells populated, 
    '        Dim LineCellsPopulated As Long = 0
    '        For DetailColumn = 1 To ws.UsedRange.Columns.Count
    '            ' Cell Has Content
    '            If Not IsNothing(ws.Cells(DetailRow, DetailColumn).value) Then
    '                If ws.Cells(DetailRow, DetailColumn).Value.ToString.Trim <> "" Then
    '                    LineCellsPopulated = LineCellsPopulated + 1
    '                End If
    '            End If
    '        Next

    '        ' We expect all lines below the first line to have at least 70% of the content of the first line - If not than merge content up and delete row
    '        If (LineCellsPopulated / FirstLineCellsPopulated) < 0.7 Then
    '            CleanedLines = True

    '            ' We need to merge current rows (with not enough data) together with row above
    '            Dim MergeToRow As Long = DetailRow - 1
    '            For DetailColumn = 1 To ws.UsedRange.Columns.Count
    '                If Not IsNothing(ws.Cells(DetailRow, DetailColumn).value) Then
    '                    If ws.Cells(DetailRow, DetailColumn).Value.ToString.Trim <> "" Then
    '                        ' Merge Cells Data
    '                        ws.Cells(MergeToRow, DetailColumn).Value = ws.Cells(MergeToRow, DetailColumn).Value + " " + ws.Cells(DetailRow, DetailColumn).Value
    '                    End If
    '                End If
    '            Next

    '            ' Delete Row
    '            ws.Rows(DetailRow).EntireRow.Delete

    '            ' As we have removed row, set index back 1 and reduce last row
    '            DetailRow = DetailRow - 1
    '            LastRow = LastRow - 1

    '            If DetailRow >= LastRow Then
    '                Exit For
    '            End If
    '        End If
    '    Next

    '    ' Remove Columns with no data (only details)
    '    For DetailColumn = ws.UsedRange.Columns.Count To 2 Step -1
    '        Dim bPreviousColumnHasData As Boolean = False

    '        For DetailRow As Long = HeaderRow + HeaderRowSpan To LastRow
    '            If Not IsNothing(ws.Cells(DetailRow, DetailColumn - 1).Value) Then
    '                If ws.Cells(DetailRow, DetailColumn - 1).Value.ToString.Trim <> "" Then bPreviousColumnHasData = True
    '            End If
    '        Next

    '        ' If Previous Column Doesn't have data - move all the data across from this point to end of array
    '        If Not bPreviousColumnHasData Then
    '            For DetailRow As Long = HeaderRow + HeaderRowSpan To LastRow
    '                For MoveFromColumn As Long = DetailColumn To ws.UsedRange.Columns.Count
    '                    Dim MoveToColumn As Long = MoveFromColumn - 1

    '                    ws.Cells(DetailRow, MoveToColumn).Value = ws.Cells(DetailRow, MoveFromColumn).Value
    '                    ws.Cells(DetailRow, MoveFromColumn).Value = ""
    '                Next

    '                CleanedLines = True
    '            Next
    '        End If
    '    Next

    '    ExcelApp.ScreenUpdating = True

    '    Return CleanedLines
    'End Function

    'Private Function GetDataFromDetails(ByRef HeaderData As Dictionary(Of String, String), ByRef DetailsData As DataTable, ProcessingData As Data.DataTable, HeaderRow As Long, HeaderRowSpan As Long, DetailsRowSpan As Long, LastRow As Long)
    '    ' Add Records for each row
    '    Dim RowID As Integer = 1
    '    For DetailRow As Long = HeaderRow + HeaderRowSpan To LastRow Step DetailsRowSpan
    '        'Dim dr As DataRow = DetailsData.NewRow
    '        'dr("ID") = ID
    '        DetailsData.Rows.Add(RowID)
    '        RowID += 1
    '    Next

    '    ' Go through each piece of data we get for details
    '    For DetailsCol As Long = 1 To 7
    '        Dim DetailHeader = ""
    '        Dim DetailColumn As Long
    '        Dim CellColumnSplit As String = ""
    '        Dim CellColumn As String = ""

    '        Select Case DetailsCol
    '            Case 1 : DetailHeader = "ItemsCode"
    '            Case 2 : DetailHeader = "ItemsDescription"
    '            Case 3 : DetailHeader = "ItemsQty"
    '            Case 4 : DetailHeader = "ItemsExGST"
    '            Case 5 : DetailHeader = "ItemsIncGST"
    '            Case 6 : DetailHeader = "ItemsExExtended"
    '            Case 7 : DetailHeader = "ItemsIncExtended"
    '        End Select

    '        ' Find the header in the header cells and get the column it is in
    '        DetailColumn = GetColForDetailFromSpreadsheet(ProcessingData, DetailHeader, HeaderRow, LastRow, CellColumnSplit, CellColumn)

    '        If DetailColumn > 0 Then
    '            ' Go through each of the rows below and update the Details Table
    '            RowID = 1
    '            For DetailRow As Long = HeaderRow + HeaderRowSpan To LastRow Step DetailsRowSpan
    '                Dim DetailValue As String = SheetContents(DetailRow, DetailColumn)

    '                If CellColumnSplit <> "" And IsNumeric(CellColumn) Then
    '                    Dim CellsColumns() As String = Split(DetailValue, CellColumnSplit.Replace("""", ""))
    '                    ' Remove any empty values from array
    '                    For i = CellsColumns.Length - 1 To 0 Step -1
    '                        If CellsColumns(i).Trim = "" Then
    '                            CellsColumns = CellsColumns.Where(Function(item, index) index <> i).ToArray
    '                        Else
    '                            CellsColumns(i) = CellsColumns(i).Trim
    '                        End If
    '                    Next

    '                    If CellsColumns.Count >= CLng(CellColumn) Then
    '                        DetailValue = CellsColumns(CLng(CellColumn) - 1) ' -1 is due to 0 index array
    '                    End If
    '                End If

    '                Select Case DetailsCol
    '                    Case 1 : DetailsData.Rows(RowID - 1)("Code") = DetailValue
    '                    Case 2 : DetailsData.Rows(RowID - 1)("Description") = DetailValue
    '                    Case 3 : If IsNumeric(DetailValue) Then DetailsData.Rows(RowID - 1)("Qty") = CDec(DetailValue)
    '                    Case 4 : If IsNumeric(DetailValue) Then DetailsData.Rows(RowID - 1)("UnitExGST") = CDec(DetailValue)
    '                    Case 5 : If IsNumeric(DetailValue) Then DetailsData.Rows(RowID - 1)("UnitIncGST") = CDec(DetailValue)
    '                    Case 6 : If IsNumeric(DetailValue) Then DetailsData.Rows(RowID - 1)("ExtendedExGST") = CDec(DetailValue)
    '                    Case 7 : If IsNumeric(DetailValue) Then DetailsData.Rows(RowID - 1)("ExtendedIncGST") = CDec(DetailValue)
    '                End Select

    '                RowID += 1
    '            Next
    '        End If
    '    Next

    '    Return True
    'End Function

    'Private Sub UnMergeCells(ws As Excel.Worksheet)
    '    Dim C As Excel.Range

    '    ExcelApp.FindFormat.Clear()
    '    ExcelApp.FindFormat.MergeCells = True
    '    C = ws.Cells.Find("", SearchFormat:=True)
    '    Do While Not C Is Nothing
    '        C.UnMerge()
    '        C = ws.Cells.Find("", SearchFormat:=True)
    '    Loop
    '    ExcelApp.FindFormat.Clear()
    'End Sub

    'Private Sub DeleteEmptyRowsAndCells(ws As Excel.Worksheet)
    '    ' Clear the Empty Cells
    '    ' For Each Row (backwards - to allow for deletes
    '    For row As Long = ws.UsedRange.Rows.Count To 1 Step -1
    '        Dim RowEmpty As Boolean = True
    '        ' Go through Columns, If empty delete the cell
    '        For col As Long = ws.UsedRange.Columns.Count - 1 To 1 Step -1
    '            Dim cell As Excel.Range = ws.Cells(row, col)
    '            Dim DeleteCell As Boolean = False

    '            If Not String.IsNullOrEmpty(cell.Value) Then
    '                If cell.Value.ToString = "" Then
    '                    DeleteCell = True
    '                Else
    '                    RowEmpty = False
    '                End If
    '            Else
    '                DeleteCell = True
    '            End If

    '            If DeleteCell Then cell.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)
    '        Next

    '        If RowEmpty Then
    '            ws.Rows(row).EntireRow.Delete
    '        End If
    '    Next
    'End Sub

    'Private Sub AddHeaderData(d As Dictionary(Of String, String), Key As String, Data As String)
    '    If d.ContainsKey(Key) Then
    '        d(Key) = Data
    '    Else
    '        d.Add(Key, Data)
    '    End If

    'End Sub

    'Private Sub SetInvoiceInvalid(d As Dictionary(Of String, String), Exception As String)
    '    If d.ContainsKey("InvoiceValid") Then
    '        d("InvoiceValid") = "FALSE"
    '    Else
    '        d.Add("InvoiceValid", "FALSE")
    '    End If

    '    If d.ContainsKey("Exceptions") Then
    '        d("Exceptions") = d("Exceptions") + Environment.NewLine + Exception
    '    Else
    '        d.Add("Exceptions", Exception)
    '    End If

    'End Sub

    'Private Function GetProcessingTemplateData(ProcessingTemplate As String) As Data.DataTable
    '    Dim wbTemplate As Excel.Workbook
    '    Dim wsTemplate As Excel.Worksheet

    '    If IsNothing(ExcelApp) Then
    '        ExcelApp = New Excel.Application
    '        ExcelApp.Visible = True
    '    End If

    '    wbTemplate = ExcelApp.Workbooks.Open(ProcessingTemplate,, True)

    '    ' Set the Worksheet
    '    wsTemplate = wbTemplate.ActiveSheet

    '    Dim dt As New Data.DataTable

    '    ' Create typed columns in the DataTable.
    '    dt.Columns.Add("Field", GetType(String))
    '    dt.Columns.Add("Comment", GetType(String))

    '    dt.Columns.Add("SearchText", GetType(String))
    '    dt.Columns.Add("SearchDirection", GetType(String))

    '    dt.Columns.Add("AnchorPosition", GetType(String))
    '    dt.Columns.Add("AnchorCellsToMove", GetType(String))
    '    dt.Columns.Add("AnchorRowStart", GetType(String))
    '    dt.Columns.Add("AnchorRowEnd", GetType(String))
    '    dt.Columns.Add("AnchorColumnStart", GetType(String))
    '    dt.Columns.Add("AnchorColumnEnd", GetType(String))

    '    dt.Columns.Add("ValueMaxChars", GetType(String))

    '    dt.Columns.Add("RowSpan", GetType(String))

    '    For DetailsCol As Long = 1 To 7
    '        Dim DetailHeader As String = ""

    '        Select Case DetailsCol
    '            Case 1 : DetailHeader = "ItemsCode"
    '            Case 2 : DetailHeader = "ItemsDescription"
    '            Case 3 : DetailHeader = "ItemsQty"
    '            Case 4 : DetailHeader = "ItemsExGST"
    '            Case 5 : DetailHeader = "ItemsIncGST"
    '            Case 6 : DetailHeader = "ItemsExExtended"
    '            Case 7 : DetailHeader = "ItemsIncExtended"
    '        End Select

    '        dt.Columns.Add(DetailHeader + "Header", GetType(String))
    '        dt.Columns.Add(DetailHeader + "SheetColumn", GetType(String))
    '        dt.Columns.Add(DetailHeader + "CellColumnSplit", GetType(String))
    '        dt.Columns.Add(DetailHeader + "CellColumn", GetType(String))
    '        dt.Columns.Add(DetailHeader + "Length", GetType(String))
    '    Next

    '    ' Get the Sheet Contents into Array for fast reading
    '    Dim SheetContents As Object
    '    SheetContents = wsTemplate.Range(wsTemplate.Cells(3, 1), wsTemplate.Cells(wsTemplate.UsedRange.Rows.Count, wsTemplate.UsedRange.Columns.Count)).Value

    '    ' Read Rows
    '    For r As Long = 1 To UBound(SheetContents)
    '        Dim dr As DataRow = dt.NewRow

    '        dr("Field") = SheetContents(r, 1)
    '        dr("Comment") = SheetContents(r, 2)

    '        dr("SearchText") = SheetContents(r, 3)
    '        dr("SearchDirection") = SheetContents(r, 4)

    '        dr("AnchorPosition") = SheetContents(r, 5)
    '        dr("AnchorCellsToMove") = SheetContents(r, 6)
    '        dr("AnchorRowStart") = SheetContents(r, 7)
    '        dr("AnchorRowEnd") = SheetContents(r, 8)
    '        dr("AnchorColumnStart") = SheetContents(r, 9)
    '        dr("AnchorColumnEnd") = SheetContents(r, 10)

    '        dr("ValueMaxChars") = SheetContents(r, 11)

    '        dr("RowSpan") = SheetContents(r, 12)

    '        ' Get the Detail Lines Settings
    '        Dim ColStart As Long = 13

    '        For DetailsCol As Long = 1 To 7
    '            Dim DetailHeader As String = ""

    '            Select Case DetailsCol
    '                Case 1 : DetailHeader = "ItemsCode"
    '                Case 2 : DetailHeader = "ItemsDescription"
    '                Case 3 : DetailHeader = "ItemsQty"
    '                Case 4 : DetailHeader = "ItemsExGST"
    '                Case 5 : DetailHeader = "ItemsIncGST"
    '                Case 6 : DetailHeader = "ItemsExExtended"
    '                Case 7 : DetailHeader = "ItemsIncExtended"
    '            End Select

    '            If DetailsCol > 1 Then ColStart = ColStart + 5

    '            dr(DetailHeader + "Header") = SheetContents(r, ColStart)
    '            dr(DetailHeader + "SheetColumn") = SheetContents(r, ColStart + 1)
    '            dr(DetailHeader + "CellColumnSplit") = SheetContents(r, ColStart + 2)
    '            dr(DetailHeader + "CellColumn") = SheetContents(r, ColStart + 3)
    '            dr(DetailHeader + "Length") = SheetContents(r, ColStart + 4)
    '        Next

    '        dt.Rows.Add(dr)
    '    Next

    '    wbTemplate.Close(False)

    '    Return dt

    'End Function

    'Private Function CreasteDetailsDataTable() As Data.DataTable
    '    Dim dt As New Data.DataTable

    '    ' Create typed columns in the DataTable.
    '    dt.Columns.Add("ID", GetType(Integer))
    '    dt.Columns.Add("Code", GetType(String))
    '    dt.Columns.Add("Description", GetType(String))
    '    dt.Columns.Add("Qty", GetType(Decimal))
    '    dt.Columns.Add("UnitExGST", GetType(Decimal))
    '    dt.Columns.Add("UnitIncGST", GetType(Decimal))
    '    dt.Columns.Add("ExtendedExGST", GetType(Decimal))
    '    dt.Columns.Add("ExtendedIncGST", GetType(Decimal))

    '    Return dt

    'End Function

    'Private Function GetCompanyName(ABN As String) As String
    '    Dim Search As httpXMLSearch
    '    Dim SearchPayload As String
    '    Dim MainName As String = ""
    '    Dim MainTradingName As String = ""
    '    Dim CompanyName As String = ""

    '    Log(LogLevels.Trace, "GetCompanyName - Declaring Search Variable...")
    '    Search = New httpXMLDocumentSearch

    '    Log(LogLevels.Trace, "GetCompanyName - Performing ABN Search...")
    '    SearchPayload = Search.ABNSearch(ABN, "n", "371f387c-0a3b-420b-ac39-00bb04f5b85f")

    '    Log(LogLevels.Trace, "GetCompanyName - Parsing Result...")
    '    Dim p As New XMLMessageParser(SearchPayload)

    '    Return p.GetCompanyName()

    'End Function

    'Private Function GetValueFromSpreadsheet(Field As String, FieldType As FieldTypes, TemplateData As Data.DataTable, Optional ByRef RowSpan As Long = 0) As String
    '    Dim sResult As String = ""
    '    Dim lRow As Long = 0

    '    ' Filter DataTable to Field
    '    Dim rows As Data.DataRow() = TemplateData.Select("Field = '" + Field + "'")

    '    For Each row As DataRow In rows
    '        GetContentForFieldUsingAnchor(Field, FieldType, row, sResult, lRow)
    '        If sResult <> "" Then
    '            ' For Details Header/Lines get if they span multiple Rows
    '            If IsNumeric(row("RowSpan").ToString) Then
    '                RowSpan = row("RowSpan").ToString
    '            Else
    '                RowSpan = 1
    '            End If

    '            Exit For
    '        End If
    '    Next

    '    Return sResult
    'End Function

    'Private Function GetRowFromSpreadsheet(Field As String, FieldType As FieldTypes, TemplateData As Data.DataTable, Optional ByRef RowSpan As Long = 0) As Long
    '    Dim sResult As String = ""
    '    Dim lRow As Long = 0

    '    ' Filter DataTable to Field
    '    Dim rows As Data.DataRow() = TemplateData.Select("Field = '" + Field + "'")

    '    For Each row As DataRow In rows
    '        FindAnchor(Field, FieldType, row, sResult, lRow)

    '        If lRow > 0 Then
    '            ' For Details Header/Lines get if they span multiple Rows
    '            If IsNumeric(row("RowSpan").ToString) Then
    '                RowSpan = row("RowSpan").ToString
    '            Else
    '                RowSpan = 1
    '            End If

    '            Exit For
    '        End If
    '    Next

    '    Return lRow
    'End Function

    'Private Function GetColForDetailFromSpreadsheet(TemplateData As Data.DataTable, DetailHeader As String, RowStart As Long, RowEnd As Long, ByRef CellColumnSplit As String, ByRef CellColumn As String) As Long
    '    Dim sResult As String = ""
    '    Dim lRow As Long = 0
    '    Dim lCol As Long = 0

    '    ' Filter DataTable to Field
    '    Dim rows As Data.DataRow() = TemplateData.Select("Field = 'InvoiceDetails'")

    '    For Each row As DataRow In rows

    '        ' **** Find the Column holding the Data ****
    '        Dim MaxLength As String = row(DetailHeader + "Length").ToString

    '        ' If Column is specified in Template return this
    '        If IsNumeric(row(DetailHeader + "SheetColumn").ToString) Then
    '            lCol = CLng(row(DetailHeader + "SheetColumn").ToString)

    '        ElseIf row(DetailHeader + "Header").ToString <> "" Then ' Search for it in the Header Rows
    '            GetUsingAnchor(False, FieldTypes.String, row(DetailHeader + "Header").ToString, "", "", "", RowStart, RowEnd, "", "", MaxLength, sResult, lRow, lCol)
    '        End If

    '        ' Get the Cell Column Split Content
    '        CellColumnSplit = row(DetailHeader + "CellColumnSplit").ToString
    '        CellColumn = row(DetailHeader + "CellColumn").ToString
    '    Next

    '    Return lCol
    'End Function

    'Private Sub GetContentForFieldUsingAnchor(Field As String, FieldType As FieldTypes, row As Data.DataRow, ByRef ResultValue As String, Optional ByRef ResultRow As Long = 0, Optional ByRef ResultColumn As Long = 0)

    '    GetUsingAnchor(True,
    '                   FieldType,
    '                   row("SearchText").ToString,
    '                   row("SearchDirection").ToString,
    '                   row("AnchorPosition").ToString,
    '                   row("AnchorCellsToMove").ToString,
    '                   row("AnchorRowStart").ToString,
    '                   row("AnchorRowEnd").ToString,
    '                   row("AnchorColumnStart").ToString,
    '                   row("AnchorColumnEnd").ToString,
    '                   row("ValueMaxChars").ToString,
    '                   ResultValue,
    '                   ResultRow,
    '                   ResultColumn)

    '    If ResultValue <> "" Then
    '        Debug.Print("Found " + Field + " Value in Row " + (ResultRow).ToString + " Column " + (ResultColumn).ToString + " = " + ResultValue)
    '    End If

    'End Sub

    'Private Sub FindAnchor(Field As String, FieldType As FieldTypes, row As Data.DataRow, ByRef ResultValue As String, Optional ByRef ResultRow As Long = 0, Optional ByRef ResultColumn As Long = 0)

    '    GetUsingAnchor(False,
    '                   FieldType,
    '                   row("SearchText").ToString,
    '                   row("SearchDirection").ToString,
    '                   row("AnchorPosition").ToString,
    '                   row("AnchorCellsToMove").ToString,
    '                   row("AnchorRowStart").ToString,
    '                   row("AnchorRowEnd").ToString,
    '                   row("AnchorColumnStart").ToString,
    '                   row("AnchorColumnEnd").ToString,
    '                   row("ValueMaxChars").ToString,
    '                   ResultValue,
    '                   ResultRow,
    '                   ResultColumn)

    '    If ResultValue <> "" Then
    '        Debug.Print("Found " + Field + " in Row " + (ResultRow).ToString + " Column " + (ResultColumn).ToString)
    '    End If

    'End Sub

    'Private Sub GetUsingAnchor(GetData As Boolean, FieldType As FieldTypes, SearchText As String, SearchDirection As String, AnchorPosition As String, AnchorCellsToMove As String, RowStart As String, RowEnd As String, ColStart As String, ColEnd As String, ValueMaxChars As String, ByRef ResultValue As String, Optional ByRef ResultRow As Long = 0, Optional ByRef ResultColumn As Long = 0)
    '    Dim sInput As String
    '    Dim sResult As String = ""
    '    Dim SearchPattern As String = ""
    '    Dim r As Long
    '    Dim c As Long
    '    Dim RowsAdded As Long
    '    Dim ColumnsAdded As Long

    '    ' Get the Defulat Row & Col Start, End & Step Settings (Direction - Top Down)
    '    Dim rStart As Long = 1
    '    Dim rEnd As Long = UBound(SheetContents, 1)
    '    Dim rStep As Long = 1
    '    Dim cStart As Long = 1
    '    Dim cEnd As Long = UBound(SheetContents, 2)
    '    Dim cStep As Long = 1

    '    ' Get the Search Direction
    '    If SearchDirection.ToUpper = "BOTTOM UP" Then
    '        rStart = UBound(SheetContents, 1)
    '        rEnd = 1
    '        rStep = -1
    '        cStart = UBound(SheetContents, 2)
    '        cEnd = 1
    '        cStep = -1
    '    End If
    '    ' Get the Search Direction
    '    If IsNumeric(RowStart) Then rStart = RowStart
    '    If IsNumeric(RowEnd) Then rEnd = RowEnd
    '    If IsNumeric(ColStart) Then cStart = ColStart
    '    If IsNumeric(ColEnd) Then cEnd = ColEnd

    '    For r = rStart To rEnd Step rStep
    '        For c = cStart To cEnd Step cStep
    '            RowsAdded = 0
    '            ColumnsAdded = 0

    '            SearchPattern = SearchText

    '            If SearchPattern <> "" Then
    '                sInput = SheetContents(r, c)

    '                If sInput <> "" Then

    '                    Dim regex As Regex = New Regex(SearchPattern, RegexOptions.IgnoreCase Or RegexOptions.Multiline)
    '                    Dim result As MatchCollection = regex.Matches(sInput)

    '                    If result.Count > 0 Then
    '                        If GetData Then
    '                            sResult = GetDataForField(sInput, SheetContents, AnchorPosition, AnchorCellsToMove, ValueMaxChars, result, FieldType, r, c, RowsAdded, ColumnsAdded)
    '                        Else
    '                            sResult = GetField(sInput, result)
    '                        End If

    '                        If sResult <> "" Then
    '                            ResultRow = r + RowsAdded
    '                            ResultColumn = c + ColumnsAdded
    '                        End If
    '                    End If
    '                End If
    '            End If

    '            ' Check if we have found result - exit loop 
    '            If sResult <> "" Then Exit For
    '        Next

    '        ' Check if we have found result - exit loop 
    '        If sResult <> "" Then Exit For
    '    Next

    '    ResultValue = sResult

    'End Sub

    'Private Function GetDataForField(Input As String, SheetContents As Object, AnchorPosition As String, AnchorCellsToMove As String, ValueMaxChars As String, Result As MatchCollection, FieldType As FieldTypes, r As Long, c As Long, ByRef RowsAdded As Long, ByRef ColumnsAdded As Long) As String
    '    Dim sResult As String = ""
    '    Dim ResultLength As Long = 0
    '    If IsNumeric(ValueMaxChars) Then ResultLength = CLng(ValueMaxChars)

    '    ' Get the details after the search string
    '    ' Get the remaining text and trim
    '    sResult = Mid(Input, Result(0).Index + Result(0).Value.Length + 1).Trim

    '    ' If we have text after the identifier
    '    If InStr(sResult, vbCr) - 1 > Len(sResult) Then
    '        ' Remove everything past the eol
    '        sResult = Strings.Left(sResult, InStr(sResult, vbCr) - 1)
    '    End If

    '    ' Limit to Ascii Non Control Characters (i.e. remove tab, CR, LF, etc)
    '    sResult = Regex.Replace(sResult, "[^\x20-\x7E]", "")

    '    ' See if we have a specified number of cells to move
    '    Dim CellsToMove As Long = 0
    '    If IsNumeric(AnchorCellsToMove) Then CellsToMove = CLng(AnchorCellsToMove)

    '    Dim CellsToSearchRight As Integer = 1
    '    If AnchorPosition.ToUpper = "LEFT" Then
    '        CellsToSearchRight = 5
    '    End If

    '    ' If Empty or we need to search cells --> see if cells to right has value
    '    If sResult = "" Or CellsToSearchRight > 1 Then
    '        Dim StartCol As Long = 1

    '        If CellsToMove > 0 Then
    '            sResult = GetValueByCellsToMoveForColumn(r, c + 1, CellsToMove, ColumnsAdded)
    '        End If

    '        If sResult = "" Then
    '            ' Check Cells to right
    '            For ColumnsAdded = StartCol To CellsToSearchRight
    '                If Not IsNothing(SheetContents(r, c + ColumnsAdded)) Then
    '                    sResult = SheetContents(r, c + ColumnsAdded).ToString
    '                    If sResult <> "" Then Exit For
    '                End If
    '            Next
    '        End If
    '    End If

    '    Dim CellsToSearchDown As Integer = 1
    '    If AnchorPosition.ToUpper = "ABOVE" Then
    '        CellsToSearchDown = 5
    '    End If

    '    ' If Empty or we need to search cells --> see if cells below has value
    '    If sResult = "" Or CellsToSearchDown > 1 Then
    '        ' If user has specified CellsToMove (Cells that have values)
    '        If CellsToMove > 0 Then
    '            sResult = GetValueByCellsToMoveForRow(r + 1, c, CellsToMove, RowsAdded)
    '        Else
    '            ' Check Cell below
    '            For RowsAdded = 1 To CellsToSearchDown
    '                If Not IsNothing(SheetContents(r + RowsAdded, c)) Then
    '                    sResult = SheetContents(r + RowsAdded, c).ToString
    '                    If sResult <> "" Then Exit For
    '                Else ' Only Search Cells with values
    '                    RowsAdded = RowsAdded - 1
    '                End If
    '            Next
    '        End If
    '    End If

    '    ' If we are specifying the number of characters, only get these
    '    If ResultLength > 0 Then
    '        sResult = Strings.Left(sResult, ResultLength)
    '    End If

    '    ' Trim Result
    '    sResult = sResult.Trim

    '    ' Limit to Ascii Non Control Characters (i.e. remove tab, CR, LF, etc)
    '    sResult = Regex.Replace(sResult, "[^\x20-\x7E]", "")

    '    ' Validate & Only Extract Required Data depending on Field Type
    '    Select Case FieldType
    '        Case FieldTypes.ABN  ' Only get numerical and decimal point characters and make sure it is 11 characters
    '            sResult = Regex.Replace(sResult, "[^0-9]", "").Trim
    '            If sResult.Length <> 11 Then sResult = ""

    '        Case FieldTypes.Currency  ' Only get numerical and decimal point characters
    '            sResult = Regex.Replace(sResult, "[^0-9.]", "").Trim

    '        Case FieldTypes.Date ' Only get numerical and - or / characters and Months 
    '            Dim regexValue As Regex = New Regex("\d{1,2}(/|-)(\d{1,2}|Jan|January|Feb|February|Mar|March|Apr|April|May|Jun|June|Jul|July|Aug|August|Sep|Septembet|Oct|October|Nov|November|Dec|December)\1(\d{4}|\d{2})", RegexOptions.IgnoreCase Or RegexOptions.Multiline)
    '            Result = regexValue.Matches(sResult)

    '            If Result.Count > 0 Then
    '                sResult = Result(0).Value
    '            Else
    '                sResult = ""
    '            End If

    '    End Select

    '    Return sResult
    'End Function

    'Private Function GetValueByCellsToMoveForColumn(Row As Long, Column As Long, CellsToMove As Long, ByRef ColumnsAdded As Long) As String
    '    Dim ValuesFound As Long = 0
    '    Dim ColumnCount As Long = 0
    '    Dim Value As String = ""

    '    ' Go through each col to end 
    '    For Column = Column To UBound(SheetContents, 2)
    '        ColumnCount += 1

    '        ' If a value up the ValuesFound counter
    '        If Not IsNothing(SheetContents(Row, Column)) Then
    '            If SheetContents(Row, Column).ToString.Trim <> "" Then ValuesFound += 1
    '        End If

    '        If ValuesFound = CellsToMove Then
    '            Value = SheetContents(Row, Column)
    '            ColumnsAdded = ColumnCount
    '            Exit For
    '        End If
    '    Next

    '    Return Value
    'End Function

    'Private Function GetValueByCellsToMoveForRow(Row As Long, Column As Long, CellsToMove As Long, ByRef RowsAdded As Long) As String
    '    Dim ValuesFound As Long = 0
    '    Dim RowCount As Long = 0
    '    Dim Value As String = ""

    '    ' Go through each row 
    '    For Row = Row To UBound(SheetContents, 1)
    '        RowCount += 1

    '        ' If a value up the ValuesFound counter
    '        If Not IsNothing(SheetContents(Row, Column)) Then
    '            If SheetContents(Row, Column).ToString.Trim <> "" Then ValuesFound += 1
    '        End If

    '        If ValuesFound = CellsToMove Then
    '            Value = SheetContents(Row, Column)
    '            RowsAdded = RowCount
    '            Exit For
    '        End If
    '    Next

    '    Return Value
    'End Function

    'Private Function GetField(Input As String, Result As MatchCollection) As String
    '    Dim sResult As String = ""
    '    Dim ResultLength As Long = 0

    '    ' Get the details after the search string
    '    ' Get the remaining text and trim
    '    sResult = Input.Trim

    '    ' If we have text after the identifier
    '    If InStr(sResult, vbCr) - 1 > Len(sResult) Then
    '        ' Remove everything past the eol
    '        sResult = Strings.Left(sResult, InStr(sResult, vbCr) - 1)
    '    End If

    '    ' Limit to Ascii Non Control Characters (i.e. remove tab, CR, LF, etc)
    '    sResult = Regex.Replace(sResult, "[^\x20-\x7E]", "")

    '    Return sResult
    'End Function

End Module
