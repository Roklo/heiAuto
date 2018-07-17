
Option Explicit On
Imports Microsoft.Office.Interop.Excel
Imports GraphicsBuilder

Module modGenieManipulation

    Private m_IOList As IO_List
    Private mCitectData As clsCITECTData
    Private m_ReplaceRules As clsReplaceRules

    Dim GraphicsBuilder As GraphicsBuilderClass = New GraphicsBuilderClass()


    Private Sub ParsePage()

        mCitectData = New clsCITECTData
        mCitectData.ParsePage()

    End Sub



    Public Sub ReplaceGenieParameters(piRulesSheet As Worksheet)

        If ConfirmDangerousOperation("You are about to update ALL genies on the open citect page. Are you SURE you want to do this?") = False Then
            Exit Sub
        End If

        modConfiguration.LoadConfiguration()

        m_IOList = New IO_List
        m_IOList.ReadIOList("")

        ParsePage()

        m_ReplaceRules = New clsReplaceRules
        m_ReplaceRules.LoadRules(piRulesSheet)

        m_ReplaceRules.Execute(m_IOList, mCitectData)

        m_IOList = Nothing
        mCitectData = Nothing
        m_ReplaceRules = Nothing



    End Sub

    Public Sub FetchTagsFromIOList(piRulesSheet As Worksheet)
        If ConfirmDangerousOperation("You are about to replace all data on this sheet. Are you SURE you want to do this?") = False Then
            Exit Sub
        End If

        modConfiguration.LoadConfiguration()

        Dim lSystemstring As String

        lSystemstring = utilities.GetUserInput("Enter the systems you want to fetch separated with commas")
        If lSystemstring = "" Then
            Exit Sub
        End If

        m_IOList = New IO_List
        m_IOList.ReadIOList("")



        Dim lSystems As Collection
        lSystems = utilities.SplitStringByComma(lSystemstring)

        Dim lRow As Long
        lRow = 2
        Dim i As Long

        If lSystemstring = "*" Then
            For i = 1 To m_IOList.SystemsCount
                FillRulesSheet(m_IOList.GetSystemsByIndex(i).SYSTEM_NAME, piRulesSheet, lRow)
            Next
        Else
            For i = 1 To lSystems.Count
                lSystemstring = lSystems(i)
                FillRulesSheet(lSystemstring, piRulesSheet, lRow)
            Next
        End If

        m_IOList = Nothing
    End Sub

    Private Sub FillRulesSheet(piSystemstring As String, piRulesSheet As Worksheet, ByRef pioRow As Long)
        Dim lSys As IO_List_System
        lSys = m_IOList.GetSystemByName(piSystemstring)

        If lSys Is Nothing Then
            Exit Sub
        End If

        Dim i As Long
        Dim lPropertiesToUpdate

        For i = 1 To lSys.SignalCount
            lPropertiesToUpdate = ""
            Dim lSig As IO_List_Signal
            lSig = lSys.GetSignal(i)
            piRulesSheet.Cells(pioRow, "A") = lSig.SFI_NUMBER.SFI
            piRulesSheet.Cells(pioRow, "B") = piSystemstring
            piRulesSheet.Cells(pioRow, "C") = lSig.Description
            piRulesSheet.Cells(pioRow, "D") = lSig.SignalType
            piRulesSheet.Cells(pioRow, "E") = lSig.MODBUS_ADDRESS
            piRulesSheet.Cells(pioRow, "F") = lSig.Unit
            piRulesSheet.Cells(pioRow, "G") = lSig.IASTagname
            piRulesSheet.Cells(pioRow, "H") = ""
            If lSig.MODBUS_ADDRESS = "" Then
                piRulesSheet.Cells(pioRow, "I") = "SFI_NUMBER"
            Else
                piRulesSheet.Cells(pioRow, "I") = "MODBUS_ADDRESS"
            End If

            If LCase(lSig.Alarm_Yes_No) = "yes" Or LCase(lSig.Alarm_Yes_No) = "mon" Then
                lPropertiesToUpdate = "AlarmNo"
            Else
                lPropertiesToUpdate = "Tag"
            End If
            piRulesSheet.Cells(pioRow, "J") = lPropertiesToUpdate

            piRulesSheet.Cells(pioRow, "K") = lSig.Label
            piRulesSheet.Cells(pioRow, "L") = lSig.ToolTip

            pioRow = pioRow + 1
        Next

    End Sub





    Private Sub DO_NOT_USE_UpdateTagsFromIOList_DO_NOT_USE(piRulesSheet As Worksheet)
        'DO NOT USE!!!
        If ConfirmDangerousOperation("You are about to update all data on this sheet. Are you SURE you want to do this?") = False Then
            Exit Sub
        End If

        modConfiguration.LoadConfiguration()

        m_ReplaceRules = New clsReplaceRules
        m_ReplaceRules.LoadRules(piRulesSheet)

        m_IOList = New IO_List
        m_IOList.ReadIOList("")

        'm_ReplaceRules.UpdateFromIOList(m_IOList)


        m_IOList = Nothing
        m_ReplaceRules = Nothing
    End Sub

    Sub PasteGenies(piRulesSheet As Worksheet)

        Dim i As Long
        i = 2
        Dim lRunHrsIndex_Contr01 As Long
        Dim lRunHrsIndex_Contr02 As Long
        lRunHrsIndex_Contr01 = 1
        lRunHrsIndex_Contr02 = 1

        Dim X_Pos As Long
        Dim Y_Pos As Long
        Dim lX_Extent As Long
        Dim lY_Extent As Long


        X_Pos = 100
        Y_Pos = 40

        m_IOList = New IO_List
        m_IOList.ReadIOList("")

        Dim TAG_NAME As String
        Dim SYST_NAME As String
        Dim SIGNAL_DESCRIPTION As String
        Dim SIGNAL_DESCRIPTION1 As String
        Dim SIGNAL_TYPE As String
        Dim Modbus_adr As String
        Dim Unit As String
        Dim lTag As String
        Dim lLabel As String
        Dim Genie As String
        Dim lNode As String

        Dim lRunHrsInitFile As clsFile


        Dim lProject As String

        lProject = GraphicsBuilder.ProjectSelected

        Dim lCitectLibrary As clsCitectLibrary
        lCitectLibrary = New clsCitectLibrary

        If MsgBox("Paste alarms into current page in : " + lProject + " ?", vbOKCancel) = vbCancel Then Exit Sub


        While Not piRulesSheet.Cells(i, "B") = ""




            TAG_NAME = piRulesSheet.Cells(i, "A")
            SYST_NAME = piRulesSheet.Cells(i, "B")
            SIGNAL_DESCRIPTION = piRulesSheet.Cells(i, "C")
            SIGNAL_DESCRIPTION1 = ""
            SIGNAL_TYPE = piRulesSheet.Cells(i, "D")
            Modbus_adr = piRulesSheet.Cells(i, "E")
            Unit = piRulesSheet.Cells(i, "F")
            lTag = piRulesSheet.Cells(i, "G")
            lLabel = piRulesSheet.Cells(i, "K")
            Genie = LCase(piRulesSheet.Cells(i, "H"))

            Dim lSig As IO_List_Signal
            lSig = m_IOList.GetSignalBy(SYST_NAME, "SFI_NUMBER", TAG_NAME)

            Dim lupdate As Boolean
            lupdate = False
            Dim LDone As Boolean

            LDone = False

            Select Case Genie
            'Running hours need some special attention...
                Case "runninghours"
                    Dim lInitString As String
                    lInitString = ""
                    If InStr(LCase(lSig.NODE), "contr01") > 0 Then
                        LDone = lCitectLibrary.PasteGenie(lProject, Genie, X_Pos, Y_Pos, lSig, lX_Extent, lY_Extent)
                        If LDone Then
                            LibraryObjectPutProperty("PosNo.", CStr(lRunHrsIndex_Contr01))
                            lInitString = " PS" + "_HOUR" + CStr(lRunHrsIndex_Contr01) + "_DinamicRUN=" + lSig.Address2 + ";"
                            lRunHrsIndex_Contr01 = lRunHrsIndex_Contr01 + 1
                        End If

                    ElseIf InStr(LCase(lSig.NODE), "contr02") > 0 Then
                        LDone = lCitectLibrary.PasteGenie(lProject, Genie, X_Pos, Y_Pos, lSig, lX_Extent, lY_Extent)
                        If LDone Then
                            LibraryObjectPutProperty("PosNo.", CStr(lRunHrsIndex_Contr02))
                            lInitString = " SB" + "_HOUR" + CStr(lRunHrsIndex_Contr02) + "_DinamicRUN=" + lSig.Address2 + ";"
                            lRunHrsIndex_Contr02 = lRunHrsIndex_Contr02 + 1
                        End If
                    End If

                    Y_Pos = Y_Pos + lY_Extent - 1 ' these running hours genie to overlap by one pixel

                    If lRunHrsInitFile Is Nothing Then
                        lRunHrsInitFile = New clsFile
                        lRunHrsInitFile.OpenFile("C:\RunningHoursInit.ci")
                        lRunHrsInitFile.WriteStr("FUNCTION InitRunningHours()")
                    End If

                    lRunHrsInitFile.WriteStr(lInitString)

                Case Else

                    LDone = lCitectLibrary.PasteGenie(lProject, Genie, X_Pos, Y_Pos, lSig, lX_Extent, lY_Extent)
                    Y_Pos = Y_Pos + lY_Extent

            End Select


            If LDone Then
                piRulesSheet.Cells(i, "M") = "OK"
            Else
                piRulesSheet.Cells(i, "M") = "*" + Genie + "* NOT FOUND"
            End If


            If Y_Pos > 800 Then
                Y_Pos = 100
                X_Pos = X_Pos + 150
            End If




            i = i + 1
        End While

        If Not lRunHrsInitFile Is Nothing Then
            lRunHrsInitFile.WriteStr("END")
        End If


        m_IOList = Nothing

    End Sub


    Private Sub LibraryObjectPutProperty(piPropName As String, piValue As String)
        On Error Resume Next
        GraphicsBuilder.LibraryObjectPutProperty(piPropName, piValue)
    End Sub

    Public Sub FetchGenies(piExcelSheet As Worksheet)
        If ConfirmDangerousOperation("You are about to replace all data on this sheet. Are you SURE you want to do this?") = False Then
            Exit Sub
        End If

        Dim lPage As clsCITECTData
        lPage = New clsCITECTData
        lPage.ParsePage()


        Dim lGen As clsCITECTGenie
        Dim i As Long
        Dim lProps As clsCitectPropertyNames
        Dim Lprop As clsCITECTProperty
        lProps = lPage.GetGeniePropertyNames

        For i = 1 To lProps.PropertyCount
            Lprop = lProps.GetPropertyName(i)
            Lprop.ExcelColumn = i + 4
            piExcelSheet.Cells(1, Lprop.ExcelColumn) = Lprop.Name
        Next

        For i = 1 To lPage.GenieCount
            lGen = lPage.GetGenie(i)

            'Mark every cell grey

            piExcelSheet.Range(piExcelSheet.Cells(i + 1, 5), piExcelSheet.Cells(i + 1, 5 + lProps.PropertyCount - 1)).Interior.Color = RGB(100, 100, 100)


            piExcelSheet.Cells(i + 1, "A") = lGen.GetAnimationNumber2
            piExcelSheet.Cells(i + 1, "B") = lGen.ObjectName
            piExcelSheet.Cells(i + 1, "C") = lGen.LibraryName

            Dim lPropName As clsCITECTProperty
            Dim j As Long
            For j = 1 To lGen.PropertyCount
                Lprop = lGen.GetProperty(j)
                lPropName = lProps.GetPropertyNameByName(Lprop.Name)

                If lPropName Is Nothing Then
                    MsgBox("Offa!")
                Else
                    piExcelSheet.Cells(i + 1, lPropName.ExcelColumn) = Lprop.Value
                    piExcelSheet.Range(piExcelSheet.Cells(i + 1, lPropName.ExcelColumn), piExcelSheet.Cells(i + 1, lPropName.ExcelColumn)).Interior.Color = RGB(255, 255, 255)
                End If
            Next
        Next
    End Sub



End Module
