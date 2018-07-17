Public Class frmGenGenies

    Dim GenGenies As GenerateGenies = New GenerateGenies()

    Private Sub cmdBack_Click(sender As Object, e As EventArgs) Handles cmdBack.Click
        GenGenies.CloseExcelFile()
        GenGenies.End_Excel_App()
        frmMenu.Show()
        Me.Close()
    End Sub

    Private Sub cmdSelectFile_Click(sender As Object, e As EventArgs) Handles cmdSelectFile.Click
        GenGenies.openFile()
        lblFilePath.Text = GenGenies.filePath

    End Sub

    Private Sub cmdGenGenies_Click(sender As Object, e As EventArgs) Handles cmdGenGenies.Click
        ' Generate genies, save excel file and end all processes. Show messagebox, go to menu.
        GenGenies.GenerateGenies()
        lblFilePath.Text = "...."
        GenGenies.CloseExcelFile()
        GenGenies.End_Excel_App()
        MessageBox.Show("Genies have been generated!")
        frmMenu.Show()
        Me.Close()
    End Sub


    Private Sub frmGenGenie_Closing(sender As Object, e As EventArgs) Handles MyBase.Closing
        GenGenies.CloseExcelFile()
        GenGenies.End_Excel_App()
        'MessageBox.Show("Closing")
    End Sub

End Class