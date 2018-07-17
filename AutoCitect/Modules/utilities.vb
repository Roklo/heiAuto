Option Explicit On
Imports GraphicsBuilder
Imports Microsoft.Office.Interop.Excel



Module utilities

    Dim GraphicsBuilder As GraphicsBuilderClass = New GraphicsBuilderClass()


    Private m_ExcelSheet As Worksheet
    ' Create new Application
    Dim Excel As Application = New Application()
    ' Open Excel spreadsheet
    'Dim xlWorkbook As Workbook = Excel.Workbooks.Open("")
    ' Loop over all sheets
    'For i As Integer = 0 To w.Sheets.Count



    Private m_IOList As IO_List

    Public Function AsNumber(piVal As String) As Double
        On Error Resume Next
        Dim lRetval As Double
        lRetval = CDbl(piVal)
        AsNumber = lRetval
    End Function



    Public Function findWorkBook(piPath As String) As Workbook
        On Error Resume Next

        Dim lWB As Workbook
        For Each lWB In Excel.Workbooks
            If lWB.Path + "\" + lWB.Name = piPath Then
                findWorkBook = lWB
            End If
        Next

    End Function



    Public Function SplitStringByComma(piVal As String) As Collection
        Dim result As Object

        result = Split(piVal, ",")

        Dim lRetval As Collection
        lRetval = New Collection

        Dim i As Long
        For i = 0 To UBound(result)
            lRetval.Add(result(i))
        Next

        SplitStringByComma = lRetval

    End Function


    Public Sub GetNextAlarmNo()
        modConfiguration.LoadConfiguration()


        m_IOList = New IO_List

        m_IOList.ReadIOList("") 'modConfiguration.IOList_Path, modConfiguration.IOLIST_Sheet, modConfiguration.ColumnAlloc_Description,  ColumnAlloc_AlarmYesNo, lFrm.MODBUS_ADDRESS, lFrm.SFI_NUMBER, lFrm.NODE, lFrm.AlarmNo, lFrm.Unit, lFrm.System, lFrm.IASTagname, lFrm.ToolTip, lFrm.Label

        Dim lMaxAlmPS As String
        Dim lMaxAlmSB As String

        lMaxAlmPS = m_IOList.GetMaxAlarmPS()
        lMaxAlmSB = m_IOList.GetMaxAlarmSB()


        MsgBox("MAX ALM PS = " + CStr(lMaxAlmPS) + ", MAX ALM SB = " + CStr(lMaxAlmSB))


        m_IOList = Nothing

    End Sub

    Public Function CurrentProject() As String
        CurrentProject = GraphicsBuilder.ProjectSelected
    End Function

    Public Function ConfirmDangerousOperation(piDescription As String) As Boolean
        ConfirmDangerousOperation = False
        If MsgBox(piDescription, vbOKCancel + vbExclamation) = MsgBoxResult.Cancel Then
            Exit Function
        End If

        If Not MsgBox("ARE YOU 100% sure?? " + piDescription, vbYesNo + vbExclamation) = MsgBoxResult.Yes Then
            Exit Function
        End If

        ConfirmDangerousOperation = True


    End Function

    Public Function GetUserInput(piMessageToUser As String) As String
        Dim lInputForm As frmUserInput
        lInputForm = New frmUserInput
        GetUserInput = lInputForm.ShowDialog(piMessageToUser)

    End Function

    Public Function zero_if_blank(piValue As String) As String
        If piValue = "" Then piValue = "0"
        zero_if_blank = piValue
    End Function



End Module
