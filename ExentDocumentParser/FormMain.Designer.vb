<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.WordDocument = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ExcelWorkbook = New System.Windows.Forms.TextBox()
        Me.btnProcessDocument = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.PDFDocument = New System.Windows.Forms.TextBox()
        Me.btnSearchPDF = New System.Windows.Forms.Button()
        Me.DisplayOfficeApps = New System.Windows.Forms.CheckBox()
        Me.GenerateSampleWorkbook = New System.Windows.Forms.CheckBox()
        Me.Log = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'WordDocument
        '
        Me.WordDocument.Enabled = False
        Me.WordDocument.Location = New System.Drawing.Point(85, 74)
        Me.WordDocument.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.WordDocument.Name = "WordDocument"
        Me.WordDocument.Size = New System.Drawing.Size(892, 23)
        Me.WordDocument.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 74)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 15)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Word:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 44)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 15)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Excel:"
        '
        'ExcelWorkbook
        '
        Me.ExcelWorkbook.Enabled = False
        Me.ExcelWorkbook.Location = New System.Drawing.Point(85, 44)
        Me.ExcelWorkbook.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.ExcelWorkbook.Name = "ExcelWorkbook"
        Me.ExcelWorkbook.Size = New System.Drawing.Size(892, 23)
        Me.ExcelWorkbook.TabIndex = 6
        '
        'btnProcessDocument
        '
        Me.btnProcessDocument.Location = New System.Drawing.Point(365, 168)
        Me.btnProcessDocument.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnProcessDocument.Name = "btnProcessDocument"
        Me.btnProcessDocument.Size = New System.Drawing.Size(262, 45)
        Me.btnProcessDocument.TabIndex = 10
        Me.btnProcessDocument.Text = "Process Document"
        Me.btnProcessDocument.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(15, 14)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(31, 15)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "PDF:"
        '
        'PDFDocument
        '
        Me.PDFDocument.Location = New System.Drawing.Point(85, 14)
        Me.PDFDocument.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PDFDocument.Name = "PDFDocument"
        Me.PDFDocument.Size = New System.Drawing.Size(892, 23)
        Me.PDFDocument.TabIndex = 13
        Me.PDFDocument.Text = "C:\Users\marcu\AppData\Local\Exent\DocumentParser\Samples\Crowther Operations Pty" &
    " Ltd  Tas The Boss Shop Invoice 604984.PDF"
        '
        'btnSearchPDF
        '
        Me.btnSearchPDF.Location = New System.Drawing.Point(944, 14)
        Me.btnSearchPDF.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnSearchPDF.Name = "btnSearchPDF"
        Me.btnSearchPDF.Size = New System.Drawing.Size(32, 22)
        Me.btnSearchPDF.TabIndex = 15
        Me.btnSearchPDF.Text = "..."
        Me.btnSearchPDF.UseVisualStyleBackColor = True
        '
        'DisplayOfficeApps
        '
        Me.DisplayOfficeApps.AutoSize = True
        Me.DisplayOfficeApps.Location = New System.Drawing.Point(415, 110)
        Me.DisplayOfficeApps.Name = "DisplayOfficeApps"
        Me.DisplayOfficeApps.Size = New System.Drawing.Size(175, 19)
        Me.DisplayOfficeApps.TabIndex = 16
        Me.DisplayOfficeApps.Text = "Display Office Applications"
        Me.DisplayOfficeApps.UseVisualStyleBackColor = True
        '
        'GenerateSampleWorkbook
        '
        Me.GenerateSampleWorkbook.AutoSize = True
        Me.GenerateSampleWorkbook.Location = New System.Drawing.Point(415, 135)
        Me.GenerateSampleWorkbook.Name = "GenerateSampleWorkbook"
        Me.GenerateSampleWorkbook.Size = New System.Drawing.Size(177, 19)
        Me.GenerateSampleWorkbook.TabIndex = 17
        Me.GenerateSampleWorkbook.Text = "Generate Sample Workbook"
        Me.GenerateSampleWorkbook.UseVisualStyleBackColor = True
        '
        'Log
        '
        Me.Log.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.Log.Location = New System.Drawing.Point(18, 232)
        Me.Log.Multiline = True
        Me.Log.Name = "Log"
        Me.Log.ReadOnly = True
        Me.Log.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Log.Size = New System.Drawing.Size(959, 184)
        Me.Log.TabIndex = 18
        '
        'FormMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(983, 428)
        Me.Controls.Add(Me.Log)
        Me.Controls.Add(Me.GenerateSampleWorkbook)
        Me.Controls.Add(Me.DisplayOfficeApps)
        Me.Controls.Add(Me.btnSearchPDF)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.PDFDocument)
        Me.Controls.Add(Me.btnProcessDocument)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ExcelWorkbook)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.WordDocument)
        Me.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Name = "FormMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Exent Document Parser"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents WordDocument As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents ExcelWorkbook As TextBox
    Friend WithEvents btnProcessDocument As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents PDFDocument As TextBox
    Friend WithEvents btnSearchPDF As Button
    Friend WithEvents DisplayOfficeApps As CheckBox
    Friend WithEvents GenerateSampleWorkbook As CheckBox
    Friend WithEvents Log As TextBox
End Class
