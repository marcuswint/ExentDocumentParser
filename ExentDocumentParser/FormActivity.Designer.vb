<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormActivity
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormActivity))
        Me.lblCurrentActivity = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblCurrentActivity
        '
        Me.lblCurrentActivity.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurrentActivity.Location = New System.Drawing.Point(12, 13)
        Me.lblCurrentActivity.Name = "lblCurrentActivity"
        Me.lblCurrentActivity.Size = New System.Drawing.Size(452, 21)
        Me.lblCurrentActivity.TabIndex = 0
        Me.lblCurrentActivity.Text = "Waiting..."
        '
        'FormActivity
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(483, 43)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblCurrentActivity)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FormActivity"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Exent Document Parser"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblCurrentActivity As Windows.Forms.Label
End Class
