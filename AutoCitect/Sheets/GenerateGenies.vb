Option Explicit On
Imports System.IO
Imports Microsoft.Office.Interop.Excel

Public Class GenerateGenies

    Dim dateStart As Date = Date.Now
    Dim Excel As Application = New Application()
    Private m_xlWorkBook_IOlist As Workbook
    Private m_xlWorkSheet_IOlist As Worksheet
    Private m_xlWorkBook_GenGenies As Workbook
    Private m_xlWorkSheet_GenGenies As Worksheet
    Public filePath As String = ""
    Dim clsFile As clsFile = New clsFile()
    Dim dateEnd As Date


    Dim SFI As String
    Dim ISA As String



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


    Public Sub openFile()

        Dim myStream As Stream = Nothing
        Dim openFileDialog1 As New OpenFileDialog()
        Dim CellInfo As String
        Dim UnableToOpen As Boolean = True

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
                    UnableToOpen = False
                    dateEnd = Date.Now
                End If
            Catch ex As Exception
                MessageBox.Show("Cannot read file from disk... Error: " + ex.Message)
                UnableToOpen = True
                dateEnd = Date.Now
                'End_Excel_App(dateStart, dateEnd)

            Finally
                If (myStream IsNot Nothing) Then
                    myStream.Close()
                End If
            End Try
        End If
        If Not UnableToOpen Then
            m_xlWorkBook_IOlist = Excel.Workbooks.Open(filePath)
            m_xlWorkSheet_IOlist = m_xlWorkBook_IOlist.Worksheets("sheet1")
            m_xlWorkBook_GenGenies = Excel.Workbooks.Open("C:\Work\AutoCitect\AutoCitect\AutoCitect\Excel Files\GenGenies.xls")
            m_xlWorkSheet_GenGenies = m_xlWorkBook_GenGenies.Worksheets("sheet1")



            ' Displays the cell value B2
            'CellInfo = m_xlWorkSheet_IOlist.Cells(1, 1).value.ToString
            'MsgBox("Cell 2B: " + CellInfo)
            ' Edit the cell with new value
            'm_xlWorkSheet.Cells(2, 2) = "Hade Robin"
            'ClearPage()

        End If


        'm_xlWorkBook.Close()
    End Sub

    Public Sub GenerateGenies()
        'Dim intRowCount As Integer = m_xlWorkSheet_GenGenies.[A1].End(XlDirection.xlDown).Row
        'm_xlWorkSheet_GenGenies.Range("A2:Z" & intRowCount).ClearContents()
        'm_xlWorkSheet_GenGenies.Rows.Cells("2:", m_xlWorkSheet_GenGenies.Rows.Count)

        'm_xlWorkSheet_GenGenies.Range("A2:A50").Select()
        m_xlWorkSheet_GenGenies.Range(m_xlWorkSheet_GenGenies.Cells(2, 1), m_xlWorkSheet_GenGenies.Cells(10, 1)).ClearContents()
        m_xlWorkBook_GenGenies.Save()


        MsgBox("Cleared GenGenies.")

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
        While (m_xlWorkSheet_GenGenies.Cells(i, 2) <> "")
            m_xlWorkSheet_GenGenies.Range(m_xlWorkSheet_GenGenies.Cells(i, 1), m_xlWorkSheet_GenGenies.Cells(i, 13)).ClearContents()
            i = i + 1
        End While
ERROR_HANDLER:
        Exit Sub
    End Sub

    Private Sub ClearStatus()
        Dim i As Long
        i = 2
        While m_xlWorkSheet_GenGenies.Cells(i, "M") <> ""
            m_xlWorkSheet_GenGenies.Cells(i, "m") = ""
            i = i + 1
        End While
    End Sub

    Public Sub CloseExcelFile()
        Try
            'First try this
            m_xlWorkSheet_GenGenies.Saved = True
            m_xlWorkSheet_GenGenies.Close()
            m_xlWorkSheet_IOlist.Saved = True
            m_xlWorkSheet_IOlist.Close()
            Excel.Quit()
            releaseObject(Excel)
            releaseObject(m_xlWorkBook_IOlist)
            releaseObject(m_xlWorkSheet_GenGenies)
            releaseObject(m_xlWorkBook_GenGenies)
            releaseObject(m_xlWorkSheet_IOlist)

            'Then this
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Excel)
            Excel = Nothing

        Catch ex As Exception
            Exit Try
        End Try
    End Sub

    Public Sub End_Excel_App()
        Dim xlp() As Process = Process.GetProcessesByName("EXCEL")
        For Each Process As Process In xlp
            If Process.StartTime >= dateStart And Process.StartTime <= dateEnd Then
                Process.Kill()
                Exit For
            End If
        Next
    End Sub

End Class
