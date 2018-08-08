Imports Microsoft.Office.Interop

Module ConvertWordToExcel
    Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Integer) As Integer

    Public Sub ConvertDOCXToXLSX(f As FormActivity, WordDocument As String, ExcelApp As Excel.Application, wb As Excel.Workbook, DisplayOfficeApps As Boolean)
        Dim WordApp As Word.Application
        Dim doc As Word.Document
        Dim NextRow As Long = 1
        Dim ws As Excel.Worksheet

        UpdateStatus(f, "Loading Word...")
        WordApp = New Word.Application
        WordApp.Visible = DisplayOfficeApps

        ' Set the Excel Worksheet
        ws = wb.ActiveSheet

        ' Open Word Document
        doc = WordApp.Documents.Open(WordDocument,, True)

        ' Convert any Inline Shapes (Images etc)
        For Each oCtlInlineShape As Word.InlineShape In doc.InlineShapes
            oCtlInlineShape.ConvertToShape()
        Next

        ' Extract the Contents of any Text Box's to Excel Document and Delete
        UpdateStatus(f, "Getting Content From Word...")
        For ShapeIdx As Integer = doc.Shapes.Count To 1 Step -1
            Dim oCtlShape As Word.Shape = doc.Shapes(ShapeIdx)

            Select Case oCtlShape.Type
                Case Microsoft.Office.Core.MsoShapeType.msoGroup
                    For Each oCtlGroupShape As Word.Shape In oCtlShape.GroupItems
                        Select Case oCtlGroupShape.Type
                            Case Microsoft.Office.Core.MsoShapeType.msoTextBox
                                GetTextFromWordObject(WordApp, ExcelApp, ws, NextRow, oCtlShape)
                                NextRow = ws.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1
                        End Select
                    Next
                    oCtlShape.Delete()

                Case Microsoft.Office.Core.MsoShapeType.msoPicture
                    oCtlShape.Delete()

                Case Microsoft.Office.Core.MsoShapeType.msoTextBox
                    GetTextFromWordObject(WordApp, ExcelApp, ws, NextRow, oCtlShape)
                    NextRow = ws.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1
                    oCtlShape.Delete()
            End Select
        Next

        ' Extract the Contents of the Document to Excel Document
        doc.Select()
        WordApp.Selection.Copy()

        ' Select Cell & Paste to Excel
        SetForegroundWindow(ExcelApp.Hwnd)
        ws.Range(ws.Cells(NextRow, 1).address).Select()
        System.Windows.Forms.SendKeys.SendWait("^v")

        System.Threading.Thread.Sleep(1000)
        'Application.DoEvents()

        NextRow = ws.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1

        ' Clear clipboard
        System.Windows.Forms.Clipboard.Clear()

        ' Close Word (Not Saving)
        doc.Close(False)
        WordApp.Quit()

        releaseObject(WordApp)

    End Sub

    Private Sub GetTextFromWordObject(WordApp As Word.Application, ExcelApp As Excel.Application, ws As Excel.Worksheet, NextRow As Long, oCtl As Object)
        ' Paste the contents of the Text Box into Excel
        oCtl.TextFrame.TextRange.Select()
        WordApp.Selection.Copy()

        ' Select Cell & Paste to Excel
        SetForegroundWindow(ExcelApp.Hwnd)
        ws.Range(ws.Cells(NextRow, 1).address).Select()
        System.Windows.Forms.SendKeys.SendWait("^v")

        System.Threading.Thread.Sleep(1000)
    End Sub
End Module
