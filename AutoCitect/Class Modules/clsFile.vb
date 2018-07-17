Option Explicit On
Imports Microsoft.Office.Interop.Excel
Imports System
Imports System.IO
Imports System.Text


Public Class clsFile
    Private m_File As String
    Private m_FileID As Long
    Dim Excel As Application = New Application()

    Public Function OpenFile(piFilePath As String) As Boolean

        Excel.Workbooks.Open(piFilePath)
        m_File = Path.GetFileName(piFilePath)
        MessageBox.Show("You have chosen: " + m_File)
        OpenFile = True

        m_FileID = FreeFile()

        On Error GoTo ERR_HANDLER

        'Open(m_File, For Output As m_FileID)

        ' Using fs As FileStream = File.Open(Path1, FileMode.Open)


        'End Using


        Exit Function
ERR_HANDLER:
        OpenFile = False
        m_FileID = 0
    End Function

    Public Function WriteStr(piString As String) As Boolean
        WriteStr = False

        If m_FileID > 0 Then
            Write(m_FileID, piString)
            WriteStr = True
        End If

    End Function


    Public Sub CloseFile()
        On Error GoTo ERR_HANDLER
        'Close(m_FileID)
        m_FileID = 0
ERR_HANDLER:
    End Sub


    Private Sub Class_Terminate()
        CloseFile()
    End Sub

End Class
