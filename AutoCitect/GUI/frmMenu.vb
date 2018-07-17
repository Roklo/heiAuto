Public Class frmMenu
    Private Sub cmdTestPage_Click(sender As Object, e As EventArgs) Handles cmdTestPage.Click
        AutoCitect.Show()
        Me.Close()
    End Sub

    Private Sub cmdGenGenies_Click(sender As Object, e As EventArgs) Handles cmdGenGenies.Click
        frmGenGenies.Show()
        Me.Close()
    End Sub

    Private Sub frmMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub


End Class