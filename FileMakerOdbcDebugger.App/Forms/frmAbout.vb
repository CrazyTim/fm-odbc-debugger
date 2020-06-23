Public Class frmAbout

    Private Sub About_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblVersion.Text = $"Version: {Application.ProductVersion}"
    End Sub

End Class
