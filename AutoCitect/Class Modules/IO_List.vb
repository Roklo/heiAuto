Option Explicit On
Imports Microsoft.Office.Interop.Excel

Public Class IO_List
    'Public Description As String
    'Public MODBUS_ADDRESS As String
    'Public SFI_NUMBER As String
    'Public AlarmNo As String
    'Public Unit As String
    'Public System As String
    '
    'Public IASTagname As String
    'Public ToolTip As String
    'Public Label As String
    '
    'Public Alarm As String
    'Public NODE As String
    '
    Private m_Systems As Collection
    Private m_SystemsLookup As Collection
    Private m_AllSignals As Collection
    Private Excel As Application = New Application()
    Dim xlWorkbook = Nothing


    Private m_ExcelSheet As Worksheet

    Public Sub CompareIOLists(piOtherList As IO_List, Optional piCheckIAS As Boolean = False)
        Dim lSys As IO_List_System
        Dim i As Long
        Dim j As Long

        Dim lLogFile As clsFile
        lLogFile = New clsFile
        lLogFile.OpenFile("C:\CompareIoList.csv")

        For i = 1 To m_Systems.Count
            lSys = m_Systems.Item(i)
            MsgBox("****Comparing : " + lSys.SYSTEM_NAME)
            For j = 1 To lSys.SignalCount

                Dim lSig As IO_List_Signal
                lSig = lSys.GetSignal(j)

                Dim lOtherSig As IO_List_Signal
                Dim lOtherSys As IO_List_System
                'Set lOtherSig = piOtherList.GetSignalBySFI(lSig.SFI_NUMBER.SFI) UNSAFE!! Some signals have the same tagname
                lOtherSys = piOtherList.GetSystemByName(lSys.SYSTEM_NAME)

                If lOtherSys Is Nothing Then
                    MsgBox("*****FAILD TO FIND SYSTEM**** :" + lSys.SYSTEM_NAME, vbOKOnly + vbInformation, "Error")
                End If

                lOtherSig = piOtherList.GetSignalBy(lSys.SYSTEM_NAME, "SFI_NUMBER", lSig.SFI_NUMBER.SFI)

                Dim lDifferences As String

                If lOtherSig Is Nothing Then
                    lLogFile.WriteStr(lSys.SYSTEM_NAME + ";" + lSig.SFI_NUMBER.SFI + ";" + lSig.Description + "; **NOT FOUND**")
                ElseIf lSig.IsIdentical(lOtherSig, piCheckIAS, lDifferences) Then
                    lLogFile.WriteStr(lSys.SYSTEM_NAME + ";" + lSig.SFI_NUMBER.SFI + ";" + lSig.Description + "; IDENTICAL")
                Else
                    lLogFile.WriteStr(lSys.SYSTEM_NAME + ";" + lSig.SFI_NUMBER.SFI + ";" + lSig.Description + "; **DIFFERENT** ; " + lDifferences)
                End If


            Next
        Next
        MsgBox("Differences written to: C:\CompareIOList.csv", vbOKOnly + vbInformation, "Error")
    End Sub


    Public ReadOnly Property IOLIST_Sheet() As Worksheet
        Get
            IOLIST_Sheet = m_ExcelSheet
        End Get
    End Property

    Public Function GetSheet(piWorkBook As Workbook, piSheetName As String) As Worksheet
        On Error Resume Next
        GetSheet = piWorkBook.Worksheets(piSheetName)
    End Function


    Public Function ReadIOList(piPath As String)
        m_AllSignals = Nothing
        m_SystemsLookup = New Collection


        modConfiguration.LoadConfiguration

        Dim xlWorkbook As Workbook

        Dim lPath As String
        lPath = piPath
        If lPath = "" Then lPath = modConfiguration.IOList_Path


        xlWorkbook = findWorkBook(lPath)
        If xlWorkbook Is Nothing Then
            xlWorkbook = Excel.Workbooks.Open(lPath, , True)

        End If


        Dim lSheet As Worksheet
        lSheet = GetSheet(xlWorkbook, modConfiguration.IOLIST_Sheet) ' xlWorkbook.Worksheets(modConfiguration.IOLIST_Sheet)
        If lSheet Is Nothing Then
            Dim testmsg As Integer
            testmsg = MsgBox("ERROR in ReadIOList(..) : did not find worksheet :" + modConfiguration.IOLIST_Sheet,
                                 vbOKOnly + vbInformation, "Error")
        End If


        Read(lSheet)

        m_ExcelSheet = lSheet


    End Function

    Public Sub TransferSettings(piSourceIO_List As String, piSourceDestColumn As String)

        Dim lSouceList As IO_List
        lSouceList = New IO_List
        lSouceList.ReadIOList(piSourceIO_List)

        Dim lFile As clsFile
        lFile = New clsFile
        lFile.OpenFile("C:\IOLISTTransferLog.csv")
        Dim lLogString As String

        Dim i As Long

        For i = 1 To lSouceList.SignalCount
            lLogString = ""
            Dim lSourcesig As IO_List_Signal
            lSourcesig = lSouceList.getSignalByIndex(i)

            Dim lValue As String
            lValue = lSourcesig.GetColumnInExcelSheet(piSourceDestColumn)



            Dim lTargetSig As IO_List_Signal
            Dim lSFI As String

            lSFI = lSourcesig.SFI_NUMBER.SFI
            lLogString = lSFI + ";"

            lTargetSig = Me.GetSignalBySFI(lSourcesig.SFI_NUMBER.SFI)

            If lTargetSig Is Nothing Then
                lLogString = lLogString + "NOT FOUND;"
            Else
                lLogString = lLogString + "FOUND;"
                lTargetSig.SetColumnInExcelSheet(piSourceDestColumn, lValue)
            End If
            lLogString = lLogString + piSourceDestColumn + ";" + lValue + ";" + lSourcesig.Description

            If Not lTargetSig Is Nothing Then
                lLogString = lLogString + ";" + lTargetSig.Description
            End If
            lFile.WriteStr(lLogString)

            Debug.Print(i)



        Next




    End Sub

    Public Sub CheckForDuplicateTags()
        Try
            Dim lSignalCounters As Collection
            lSignalCounters = New Collection

            Dim lFile As clsFile
            lFile = New clsFile
            lFile.OpenFile("C:\IOLISTCheckDuplicateTags.csv")
            Dim lLogString As String

            Dim i As Long

            For i = 1 To Me.SignalCount
                lLogString = ""
                Dim lSig As IO_List_Signal
                lSig = Me.getSignalByIndex(i)

                Dim lSFI As String
                lSFI = lSig.SFI_NUMBER.SFI

                If IsInCollection(lSignalCounters, lSFI) Then
                    lFile.WriteStr(lSFI + ";DUPLICATE;" + lSig.Description)
                Else
                    lSignalCounters.Add(lSFI, lSFI)
                End If
            Next

            Exit Sub
        Catch ex As Exception
            'i = i
            Dim testmsg As Integer
            testmsg = MsgBox("List of duplicate tags saved to: C:\IOLISTCheckDuplicateTags.csv",
                                 vbOKOnly + vbInformation, "Error")
        End Try
    End Sub

    Private Function IsInCollection(piColl As Collection, piKey As String) As Boolean
        On Error GoTo ERR_HANDLER

        If piColl(piKey) = piColl(piKey) Then IsInCollection = True

        Exit Function
ERR_HANDLER:
        IsInCollection = False
    End Function



    Public Function Read(piWS As Worksheet)
        m_AllSignals = Nothing


        Dim lSysName As String
        Dim lSystem As IO_List_System
        Dim cur_line As Long

        cur_line = 1
        lSysName = GetSystemName(piWS, cur_line)
        While cur_line < 5000

            lSysName = GetSystemName(piWS, cur_line)
            MsgBox(lSysName)
            ProcessSystem(piWS, lSysName, cur_line)

        End While
    End Function

    Private Sub Class_Initialize()
        m_Systems = New Collection
        m_SystemsLookup = New Collection
    End Sub



    Private Function GetSystemName(piWS As Worksheet, line As Long) As String
        GetSystemName = piWS.Cells(line, modConfiguration.ColumnAlloc_System)
    End Function

    Private Sub ProcessSystem(piWS As Worksheet, piSysName As String, ByRef cur_line As Long)
        Dim lSys As IO_List_System

        If piSysName = "" Then
            cur_line = cur_line + 1
            Exit Sub
            'MsgBox "IO_List::ProcessSystem(..), got empty string for systemname"
        End If

        lSys = GetSystemByName(piSysName)

        If lSys Is Nothing Then
            lSys = New IO_List_System
            lSys.SYSTEM_NAME = piSysName
            m_Systems.Add(lSys)
            m_SystemsLookup.Add(lSys, Trim(piSysName))
        End If

        lSys.Read(piWS, cur_line, Me)

    End Sub

    Public Function GetSystemByName(piSysName As String) As IO_List_System
        On Error Resume Next
        If piSysName = "" Then Exit Function
        GetSystemByName = m_SystemsLookup(Trim(piSysName))
    End Function

    Public Function GetSignalBy(piSystem As String, PiPropertyName As String, piPropertyValue As String) As IO_List_Signal
        Dim lSys As IO_List_System
        Dim lSig As IO_List_Signal

        lSys = GetSystemByName(piSystem)
        If Not lSys Is Nothing Then
            lSig = lSys.GetSignalBy(PiPropertyName, piPropertyValue)
        End If
        GetSignalBy = lSig
    End Function

    Public Function GetSignalBySFI(piSFI As String) As IO_List_Signal
        If piSFI = "" Then Exit Function

        Dim i As Long
        Dim lSys As IO_List_System
        Dim lSig As IO_List_Signal

        For i = 1 To m_Systems.Count
            lSys = m_Systems.Item(i)
            lSig = lSys.GetSignalBy("SFI_NUMBER", piSFI)
            If Not lSig Is Nothing Then Exit For
        Next

        GetSignalBySFI = lSig

    End Function

    Public Function GetMaxAlarmPS() As Long
        GetMaxAlarmPS = GetMaxAlarmX("Port")
    End Function

    Public Function GetMaxAlarmSB() As Long
        GetMaxAlarmSB = GetMaxAlarmX("Stbd")
    End Function


    Private Function GetMaxAlarmX(piSide As String) As Long
        If m_Systems Is Nothing Then
            Exit Function
        End If

        Dim lMaxAlarmNo As Long
        lMaxAlarmNo = -1

        Dim lSys As IO_List_System
        For Each lSys In m_Systems
            Dim i As Long
            For i = 1 To lSys.SignalCount
                Dim lSig As IO_List_Signal
                lSig = lSys.GetSignal(i)

                If (LCase(lSig.Alarm_Yes_No) = "yes" Or LCase(lSig.Alarm_Yes_No) = "mon") And InStr(lSig.NODE, piSide) > 0 Then
                    lMaxAlarmNo = ConvertToNumberAndGetMaxAlarmNo(lMaxAlarmNo, lSig.IASTagname)
                End If

            Next
        Next

        GetMaxAlarmX = lMaxAlarmNo
    End Function

    Private Function ConvertToNumberAndGetMaxAlarmNo(piAlmNo1 As Long, piAlmNo2 As String) As Long
        ConvertToNumberAndGetMaxAlarmNo = piAlmNo1
        Dim lAlmNo2 As Long

        lAlmNo2 = -1



        If piAlmNo2 = "" Then Exit Function

        While Not IsNumeric(piAlmNo2) And Len(piAlmNo2) > 0
            piAlmNo2 = Right(piAlmNo2, Len(piAlmNo2) - 1)
        End While

        If IsNumeric(piAlmNo2) Then lAlmNo2 = CLng(piAlmNo2)

        If lAlmNo2 > piAlmNo1 Then
            ConvertToNumberAndGetMaxAlarmNo = lAlmNo2
        End If

    End Function


    Public ReadOnly Property SystemsCount() As Long
        Get
            SystemsCount = m_Systems.Count
        End Get
    End Property

    Public Function GetSystemsByIndex(piIndex As Long) As IO_List_System
        GetSystemsByIndex = m_Systems(piIndex)
    End Function


    Public ReadOnly Property SignalCount() As Long
        Get
            If m_AllSignals Is Nothing Then
                m_AllSignals = New Collection
                Dim i As Long
                Dim j As Long
                For i = 1 To m_Systems.Count
                    Dim lSys As IO_List_System
                    lSys = Me.GetSystemsByIndex(i)
                    For j = 1 To lSys.SignalCount
                        Dim lSig As IO_List_Signal
                        lSig = lSys.GetSignal(j)
                        m_AllSignals.Add(lSig)
                    Next
                Next
            End If
            SignalCount = m_AllSignals.Count
        End Get

    End Property

    Public Function getSignalByIndex(piIndex As Long) As IO_List_Signal
        If piIndex > Me.SignalCount Or piIndex < 1 Then
            Exit Function
        End If
        getSignalByIndex = m_AllSignals(piIndex)
    End Function

    Public Function GetAlarms(Optional piMaxAlarmNo As Long = -1)
        Dim lAlarms As clsAlarms
        lAlarms = New clsAlarms
        If piMaxAlarmNo < 0 Then
            piMaxAlarmNo = 2 * modConfiguration.MaxAlm
        End If
        lAlarms.Initialize(piMaxAlarmNo, Me)
        GetAlarms = lAlarms
    End Function


End Class
