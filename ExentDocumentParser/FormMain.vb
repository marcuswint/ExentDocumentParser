Imports Microsoft.Office.Interop
Imports System.Data
Imports Spire.Pdf

Public Class FormMain




#Region "Get Document Contents"
#End Region

    Private Sub btnProcessDocument_Click(sender As Object, e As EventArgs) Handles btnProcessDocument.Click
        Dim Header As New Dictionary(Of String, String)
        Dim Details As New DataTable

        Me.Log.Text = ""

        GetDocumentContents(Me.PDFDocument.Text, Me.ExcelWorkbook.Text, Me.WordDocument.Text, Header, Details, Me.DisplayOfficeApps.Checked, Me.Log)

        Dim InvoiceValid As Boolean = CBool(GetHeaderData(Header, "InvoiceValid", "False"))

        Dim GenerateSample As Boolean = GenerateSampleWorkbook.Checked

        If Not InvoiceValid Then
            MsgBox("Failed to Process Document - Exception: " + GetHeaderData(Header, "Exceptions", ""))
            If GenerateSample Then
                If MsgBox("Do you still want to gentrate the sample workbook?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    GenerateSample = False
                End If
            End If
        Else
            MsgBox("Sucessfully Processed Document")
        End If


        If GenerateSample Then
            ' Generate Workbook
            Dim ExcelApp As Excel.Application = New Excel.Application
            ExcelApp.Visible = True

            Dim wb As Excel.Workbook = ExcelApp.Workbooks.Add
            Dim ws As Excel.Worksheet = wb.ActiveSheet

            ' Add Header Information
            ws.Name = "Header"

            ' Sort Header Dictionary
            Dim SortedHeader = (From entry In Header Order By entry.Value Ascending).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
            Dim Row As Long = 1

            For Each entry As KeyValuePair(Of String, String) In SortedHeader
                ws.Cells(Row, 1).value = entry.Key
                ws.Cells(Row, 2).value = entry.Value
                Row += 1
            Next

            ' Add Detail Information
            ws = wb.Sheets.Add(, ws)
            ws.Name = "Details"

            ' Add Header
            Row = 1
            For Column As Integer = 1 To Details.Columns.Count
                ws.Cells(Row, Column).value = Details.Columns.Item(Column - 1).ColumnName
            Next

            ' Details/Rows
            For Each dr As DataRow In Details.Rows
                Row += 1
                For Column As Integer = 1 To Details.Columns.Count
                    ws.Cells(Row, Column).value = dr(Column - 1).ToString
                Next
            Next
        End If

    End Sub

    Private Sub btnSearchPDF_Click(sender As Object, e As EventArgs) Handles btnSearchPDF.Click
        Dim openFileDialog1 As OpenFileDialog = New OpenFileDialog()
        openFileDialog1.Filter = "PDF Documents|*.pdf"
        openFileDialog1.Title = "Select a PDF File"

        If System.IO.File.Exists(Me.PDFDocument.Text) Then openFileDialog1.InitialDirectory = System.IO.Path.GetDirectoryName(Me.PDFDocument.Text)

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Me.PDFDocument.Text = openFileDialog1.FileName
        End If
    End Sub


End Class
