
Option Explicit On
Imports Microsoft.Office.Interop.Excel

Module modIOList


    Private m_ExcelSheet As Worksheet
    ' Create new Application
    Dim Excel As Application = New Application()
    ' Open Excel spreadsheet
    Dim xlWorkbook As Workbook = Excel.Workbooks.Open("")
    ' Loop over all sheets
    'For i As Integer = 0 To w.Sheets.Count



    Public Sub CompareIOLists(piListA As String, piListB As String)
        modConfiguration.LoadConfiguration()
        Dim lListA As IO_List
        lListA = New IO_List
        lListA.ReadIOList(piListA)


        Dim lListB As IO_List
        lListB = New IO_List
        lListB.ReadIOList(piListB)

        lListA.CompareIOLists(lListB)


    End Sub

    Public Sub TransferSettings(piSourceIO_List As String, piSourceDestColumn As String)
        modConfiguration.LoadConfiguration()
        Dim lIOList As IO_List

        lIOList = New IO_List
        lIOList.ReadIOList(piSourceIO_List) 'HVOR SKAL DENNE KOMME FRA??????????????
        lIOList.TransferSettings(piSourceIO_List, piSourceDestColumn)

    End Sub
    Public Sub CheckForDuplicateTags()
        modConfiguration.LoadConfiguration()
        Dim lIOList As IO_List

        lIOList = New IO_List
        lIOList.ReadIOList("") 'HVOR SKAL DENNE KOMME FRA??????????????

        lIOList.CheckForDuplicateTags()


    End Sub





    Public Sub CompareIOListToAlarmConfiguration()
        If ConfirmDangerousOperation("This function has to be customised, and probably wont work for you. Du you wish to continue?") = False Then
            Exit Sub
        End If

        Dim lPath As String
        Dim lAlmNameColumn As String
        Dim lAlmDescColumn As String
        Dim lAlmNoColumn As String
        Dim lTypeColumn As String
        Dim lGroupColumn As String
        Dim lDelayColumn As String
        Dim lLL_LimitColumn As String
        Dim lL_LimitColumn As String

        Dim lHH_LimitColumn As String
        Dim lH_LimitColumn As String

        Dim lRaw0Column As String
        Dim lRaw1Column As String

        Dim lEng0Column As String
        Dim lEng1Column As String





        ' ******************* PUT CONFIGURATION for alarmlist HERE*******************
        lPath = "F:\Installasjon\Prosjekt\STX-ROB\740\Automation\IAS\AYAS\7. Project Tag Lists\740_Alarms_As_Built_Temp.xls"

        lAlmNameColumn = "A"
        lAlmDescColumn = "B"
        lAlmNoColumn = "C"
        lTypeColumn = "D"
        lGroupColumn = "E"
        lDelayColumn = "F"

        lHH_LimitColumn = "J"
        lH_LimitColumn = "K"

        lLL_LimitColumn = "M"
        lL_LimitColumn = "L"

        lRaw0Column = "N"
        lRaw1Column = "O"

        lEng0Column = "P"
        lEng1Column = "Q"


        '************************************************************************

        Dim lAlmName As String
        Dim lAlmDesc As String
        Dim lAlmNo As String
        Dim lType As String
        Dim lGroup As String
        Dim lDelay As String
        Dim lLL_Limit As String
        Dim lL_Limit As String

        Dim lHH_Limit As String
        Dim lH_Limit As String

        Dim lRaw0 As String
        Dim lRaw1 As String

        Dim lEng0 As String
        Dim lEng1 As String


        modConfiguration.LoadConfiguration()
        Dim lIOList As IO_List

        lIOList = New IO_List
        lIOList.ReadIOList(lPath) 'HVOR SKAL DENNE KOMME FRA??????????????


        Dim lWorkBook As Workbook
        lWorkBook = findWorkBook(lPath)
        If lWorkBook Is Nothing Then
            lWorkBook = Excel.Workbooks.Open(lPath, , True)
        End If

        Dim lSheet As Worksheet

        '****You may also need to update the index******
        lSheet = lWorkBook.Sheets(1)
        '***********************************************

        Dim i As Long

        Dim lFile As clsFile = New clsFile()
        ' lFile = lFile.OpenFile("C:\AlarmComparison.csv")
        Dim lOutStr As String
        i = 4

        While lSheet.Cells(i, lAlmNoColumn) <> ""
            lOutStr = ""
            lAlmName = lSheet.Cells(i, lAlmNameColumn)
            lAlmDesc = lSheet.Cells(i, lAlmDescColumn)
            lAlmNo = lSheet.Cells(i, lAlmNoColumn)
            lType = lSheet.Cells(i, lTypeColumn)
            lGroup = lSheet.Cells(i, lGroupColumn)
            lDelay = lSheet.Cells(i, lDelayColumn)
            lL_Limit = lSheet.Cells(i, lL_LimitColumn)
            lLL_Limit = lSheet.Cells(i, lLL_LimitColumn)

            lH_Limit = lSheet.Cells(i, lH_LimitColumn)
            lHH_Limit = lSheet.Cells(i, lHH_LimitColumn)

            lRaw0 = lSheet.Cells(i, lRaw0Column)
            lRaw1 = lSheet.Cells(i, lRaw1Column)

            lEng0 = lSheet.Cells(i, lEng0Column)
            lEng1 = lSheet.Cells(i, lEng1Column)


            Dim lSig As IO_List_Signal
            lSig = lIOList.GetSignalBySFI(lAlmName)


            If lSig Is Nothing Then
                lOutStr = lAlmName + ";" + lAlmDesc + ";NOT FOUND;" + lAlmNo + ";" + lType
            Else
                lOutStr = lAlmName + ";" + lAlmDesc + ";FOUND;" + lAlmNo + ";" + lType

                'now look for mismatches---------------
                If lType = "1" Then 'Digital alarm
                    If Not lSig.SignalTypeEnum = IO_List_Signal.SIGNAL_TYPE.SIGNAL_TYPE_DIGITAL Then
                        lOutStr = lOutStr + ";SignalType"
                    End If

                    If Not lSig.AlarmGroup = lGroup Then
                        lOutStr = lOutStr + ";Group"
                    End If

                    If Not lSig.AlarmDelay = lDelay And Not lDelay = "0" Then
                        lOutStr = lOutStr + ";AlarmDelay"
                    End If

                    If lLL_Limit = "0" Then lLL_Limit = "NC"
                    If lLL_Limit = "1" Then lLL_Limit = "NO"

                    If Not lSig.NormalValue = lLL_Limit Then
                        lOutStr = lOutStr + ";NormalValue"
                    End If

                ElseIf lType = "101" Then 'Digital monitor point
                    If Not lSig.SignalTypeEnum = IO_List_Signal.SIGNAL_TYPE.SIGNAL_TYPE_DIGITAL Then
                        lOutStr = lOutStr + ";SignalType"
                    End If

                    If Not lSig.AlarmGroup = lGroup Then
                        lOutStr = lOutStr + ";Group"
                    End If

                    If Not lSig.AlarmDelay = lDelay And Not lDelay = "0" Then
                        lOutStr = lOutStr + ";AlarmDelay"
                    End If

                    If lLL_Limit = "0" Then lLL_Limit = "NC"
                    If lLL_Limit = "1" Then lLL_Limit = "NO"

                    If Not lSig.NormalValue = lLL_Limit Then
                        lOutStr = lOutStr + ";NormalValue"
                    End If

                ElseIf lType = "2" Then 'analog alarm
                    If Not lSig.SignalTypeEnum = IO_List_Signal.SIGNAL_TYPE.SIGNAL_TYPE_ANALOG Then
                        lOutStr = lOutStr + ";SignalType"
                    End If

                    If Not lSig.AlarmGroup = lGroup Then
                        lOutStr = lOutStr + ";Group"
                    End If

                    If Not lSig.AlarmDelay = lDelay And Not lDelay = "0" Then
                        lOutStr = lOutStr + ";AlarmDelay"
                    End If

                    If Not AsNumber(lSig.Allim_HH) = AsNumber(lHH_Limit) Then
                        lOutStr = lOutStr + ";HH_Limit"
                    End If

                    If Not AsNumber(lSig.Allim_H) = AsNumber(lH_Limit) Then
                        lOutStr = lOutStr + ";H_Limit"
                    End If

                    If Not AsNumber(lSig.AllimLL) = AsNumber(lLL_Limit) Then
                        lOutStr = lOutStr + ";LL_Limit"
                    End If

                    If Not AsNumber(lSig.AllimL) = AsNumber(lL_Limit) Then
                        lOutStr = lOutStr + ";L_Limit"
                    End If

                    If Not AsNumber(lSig.RAW_Max) = AsNumber(lRaw1) Then
                        lOutStr = lOutStr + ";Raw max"
                    End If

                    If Not AsNumber(lSig.RAW_Min) = AsNumber(lRaw0) Then
                        lOutStr = lOutStr + ";Raw min"
                    End If

                    If Not AsNumber(lSig.EngMax) = AsNumber(lEng1) Then
                        lOutStr = lOutStr + ";Eng max"
                    End If

                    If Not AsNumber(lSig.EngMin) = AsNumber(lEng0) Then
                        lOutStr = lOutStr + ";Eng min"
                    End If


                End If
                '-------------------------------------\

            End If

            If Len(lOutStr) > 0 Then
                lFile.WriteStr(lOutStr)
            End If

            i = i + 1
        End While

    End Sub

End Module
