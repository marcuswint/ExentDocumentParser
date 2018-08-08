Imports System.Data

Public Class Form1




#Region "Get Document Contents"
#End Region

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' MsgBox(GetCompanyName("51835430479"))
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim Header As New Dictionary(Of String, String)
        Dim Details As New DataTable

        GetDocumentContents(Me.TextBox1.Text, Me.TextBox2.Text, Me.TextBox3.Text, Header, Details, True)

        For i = 0 To Details.Rows.Count - 1
            MsgBox(Details.Rows(i)("ID").ToString + " - " + Details.Rows(i)("Code").ToString)
        Next

        MsgBox("Completed.")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        MsgBox(IsDate(Me.TextBox4.Text))
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub
End Class
