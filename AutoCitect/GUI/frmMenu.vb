Public Class frmMenu
    Private Sub cmdTestPage_Click(sender As Object, e As EventArgs) Handles cmdTestPage.Click
        AutoCitect.Show()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        QuickEditCitect.Show()
        Me.Close()
    End Sub
End Class