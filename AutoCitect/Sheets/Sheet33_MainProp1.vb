Option Explicit On
Imports Microsoft.Office.Interop.Excel
Imports System.IO

Public Class Sheet33_MainProp1


    Dim Excel As Application = New Application()
    Private m_xlWorkSheet As Worksheet
    Private m_xlWorkBook As Workbook
    Public filePath As String = ""
    Dim clsFile As clsFile = New clsFile()

    Private Sub cmdFetchFromIOList_Click()
        '    ClearPage
        modGenieManipulation.FetchTagsFromIOList(Me)
    End Sub

    Private Sub cmdPasteGenies_Click()
        ClearStatus()
        modGenieManipulation.PasteGenies(Me)
    End Sub

    Private Sub cmdReplace_Click()
        ClearStatus()
        modGenieManipulation.ReplaceGenieParameters(Me)
    End Sub

    Public Sub OpenFile()

        Dim myStream As Stream = Nothing
        Dim openFileDialog1 As New OpenFileDialog()
        Dim CellInfo As String
        Dim dateStart As Date = Date.Now
        Dim dateEnd As Date

        openFileDialog1.InitialDirectory = "c:\"
        openFileDialog1.RestoreDirectory = False
        openFileDialog1.FilterIndex = 1
        'openFileDialog1.Filter = "Excel files (*.xls)|.xls|ExcelX files(*.xlsx)|*.xlsx|CSV Files(*.csv)|*.csv"

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                myStream = openFileDialog1.OpenFile()
                If (myStream IsNot Nothing) Then
                    filePath = openFileDialog1.FileName()
                    MessageBox.Show("File was loaded: " + filePath)
                End If
            Catch ex As Exception
                MessageBox.Show("Cannot read file from disk... Error: " + ex.Message)
            Finally
                If (myStream IsNot Nothing) Then
                    myStream.Close()
                End If
            End Try
        End If

        m_xlWorkBook = Excel.Workbooks.Open(filePath)
        m_xlWorkSheet = m_xlWorkBook.Worksheets("sheet1")
        ' Displays the cell value B2
        CellInfo = m_xlWorkSheet.Cells(2, 2).value.ToString
        MsgBox("Cell 2B: " + CellInfo)
        ' Edit the cell with new value
        'm_xlWorkSheet.Cells(2, 2) = "Hade Robin"
        'ClearPage()

        dateEnd = Date.Now
        End_Excel_App(dateStart, dateEnd)

        'm_xlWorkBook.Close()
        'releaseObject(Excel)
        'releaseObject(m_xlWorkBook)
        'releaseObject(m_xlWorkSheet)
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub ClearPage()
        Dim i As Long
        On Error GoTo ERROR_HANDLER
        i = 2
        While (m_xlWorkSheet.Cells(i, 2) <> "")
            m_xlWorkSheet.Range(m_xlWorkSheet.Cells(i, 1), m_xlWorkSheet.Cells(i, 13)).ClearContents()
            i = i + 1
        End While
ERROR_HANDLER:
        Exit Sub
    End Sub

    Private Sub ClearStatus()
        Dim i As Long
        i = 2
        While m_xlWorkSheet.Cells(i, "M") <> ""
            m_xlWorkSheet.Cells(i, "m") = ""
            i = i + 1
        End While
    End Sub

    Private Sub End_Excel_App(dateStart As Date, dateEnd As Date)
        Dim xlp() As Process = Process.GetProcessesByName("EXCEL")
        For Each Process As Process In xlp
            If Process.StartTime >= dateStart And Process.StartTime <= dateEnd Then
                Process.Kill()
                Exit For
            End If
        Next
    End Sub
End Class
