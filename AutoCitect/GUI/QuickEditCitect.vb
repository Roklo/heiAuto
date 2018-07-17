Imports GraphicsBuilder

Public Class QuickEditCitect
    Dim CheckBox4Value As Boolean
    Dim CheckBox3Value As Boolean

    Dim GraphicsBuilder As GraphicsBuilderClass = New GraphicsBuilderClass()

    Private Sub cmdBack_Click(sender As Object, e As EventArgs) Handles cmdBack.Click
        frmMenu.Show()
        Me.Close()
    End Sub

    Private Sub ProgressBar1_Click(sender As Object, e As EventArgs) Handles ProgressBar1.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Util As Sheet2 = New Sheet2()
        Util.cmdClearLineFillColorExpression_Click()
        Dim Number As Short = GraphicsBuilder.AttributeAN()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Util As Sheet2 = New Sheet2()
        Util.cmdFindUntaggedGenies_Click()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        'Dim currentLabelValue As Integer = 0
        'Dim maxLabelValue As Integer = 0
        'If maxLabelValue = 0 Then
        '    Label1.Visible = False
        'Else
        '    Label1.Visible = True
        'End If

        'Label1.Text = currentLabelValue + "/" + maxLabelValue
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim Util As Sheet2 = New Sheet2()

        If CheckBox3Value = True Then
            Util.cmdRemoveSpesificProperty("TOOLTIP")
        End If
        If CheckBox4Value = True Then
            Util.cmdRemoveSpesificProperty("AlarmNo")
        End If


    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If Not CheckBox4.ThreeState Then
            CheckBox4.ThreeState = True
            CheckBox4Value = True
        Else
            CheckBox4.ThreeState = False
            CheckBox4Value = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If Not CheckBox3.ThreeState Then
            CheckBox3.ThreeState = True
            CheckBox3Value = True
        Else
            CheckBox3.ThreeState = False
            CheckBox3Value = False
        End If
    End Sub


End Class