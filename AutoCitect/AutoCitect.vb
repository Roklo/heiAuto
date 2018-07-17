Option Explicit On
Imports GraphicsBuilder
Imports System.IO
Imports System.Windows.Forms

Public Class AutoCitect
    Dim clsFile As clsFile = New clsFile()
    Dim GraphicsBuilder As GraphicsBuilderClass = New GraphicsBuilderClass()
    Dim FilePath As String



    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        'MessageBox.Show(GraphicsBuilder.AttributeClass)
        MessageBox.Show(GraphicsBuilder.AttributeClass + ". Press OK to GO On")

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Dim Util As Sheet2 = New Sheet2()
        Util.cmdClearLineFillColorExpression_Click()

        Dim gb As GraphicsBuilderClass = New GraphicsBuilderClass()
        Dim Number As Short = GraphicsBuilder.AttributeAN()
        'MessageBox.Show(GraphicsBuilder.AttributeAN)
        'MessageBox.Show(GraphicsBuilder.AttributeAN)
        'MessageBox.Show("press OK to GO On")

    End Sub



    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form_Load()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        gb_PasteSymbol()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'gb_PageSaved(Project, Page, LastPage)
    End Sub















    Private WithEvents gb As GraphicsBuilder.GraphicsBuilder

    Public Sub Form_Load()
        gb = New GraphicsBuilder.GraphicsBuilder
        gb.LibrarySelectionHooksEnabled = True
        gb.Visible = True
    End Sub

    Public Sub gb_PasteSymbol()
        MessageBox.Show("PasteSymbol")
    End Sub

    Private Sub gb_PageSaved(ByVal Project As String, ByVal Page As String, ByVal LastPage As Boolean)
        MessageBox.Show("PageSaved: " + Project + "." + Page + "--")
    End Sub

    Private Sub AutoCitect_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub cmdLoadFile_Click(sender As Object, e As EventArgs) Handles cmdLoadFile.Click
        Dim myStream As Stream = Nothing
        Dim openFileDialog1 As New OpenFileDialog()
        Dim fileName As String

        openFileDialog1.InitialDirectory = "c:\"
        openFileDialog1.RestoreDirectory = False
        openFileDialog1.FilterIndex = 1
        'openFileDialog1.Filter = "Excel files (*.xls)|.xls|ExcelX files(*.xlsx)|*.xlsx|CSV Files(*.csv)|*.csv"

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                myStream = openFileDialog1.OpenFile()
                If (myStream IsNot Nothing) Then
                    fileName = openFileDialog1.FileName()
                    clsFile.OpenFile(fileName)
                    lblFilePath.Text = openFileDialog1.FileName()


                    MessageBox.Show("File was loaded: " + fileName)
                End If
            Catch ex As Exception
                MessageBox.Show("Cannot read file form disk... Error: " + ex.Message)
            Finally
                If (myStream IsNot Nothing) Then
                    myStream.Close()
                End If
            End Try


        End If
    End Sub

    Private Sub cmdFilePath_TextChanged(sender As Object, e As EventArgs) Handles cmdFilePath.TextChanged, TextBox2.TextChanged
        cmdFilePath.Text = FilePath
    End Sub

    Private Sub cmdTestExcel_Click(sender As Object, e As EventArgs) Handles cmdTestExcel.Click
        Dim TestPath As String = ""
        Dim ExcelTester As Sheet33_MainProp1 = New Sheet33_MainProp1()
        ExcelTester.OpenFile()
        lblTestPath.Text = ExcelTester.filePath
    End Sub

    Private Sub cmdBack_Click(sender As Object, e As EventArgs) Handles cmdBack.Click
        frmMenu.Show()
        Me.Close()
    End Sub
End Class
