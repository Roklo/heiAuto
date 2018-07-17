Option Explicit On

Public Class frmUserInput

    Private m_TXT As String

    Public Overloads Function ShowDialog(piMessageToUser As String) As String
        lblUserInfo.Text = piMessageToUser
        Me.Show()
        ShowDialog = m_TXT
    End Function

    Private Sub cmdCancel_Click()
        Me.Hide()
    End Sub

    Private Sub cmdOK_Click()
        m_TXT = txtBox.text
        Me.Hide()
    End Sub


End Class