Option Explicit On
Imports Microsoft.VisualBasic
Imports GraphicsBuilder



Public Class Sheet2


    Private m_IOList As IO_List

    Private Sub ProcessPage_LineColor()
        'IOLIST_READ


        Dim lStat As Boolean
        Dim Ldata As clsCITECTData = New clsCITECTData
        Ldata.ParsePage()

        Dim i As Long
        Dim j As Long
        Dim LResult As DialogResult
        LResult = vbOK

        Dim a As Integer
        Dim b As Integer
        Dim c As Integer

        For i = 1 To Ldata.ObjectCount
            Dim lObj As clsCITECTObject
            lObj = Ldata.GetObject(i)

            If LCase(lObj.GetCitectObjectType) = "line" Then
                Debug.Print(lObj.GetLineColor)
                lObj.SetLineColor(209)
            End If

            If LCase(lObj.GetCitectObjectType) = "pipe" Then
                lObj.SetPipeColor(209, 209)
            End If

            If lObj.GetCitectObjectType = "CircleV2" Then
                lObj.SetLineColor(209)
            End If

            If LResult = vbCancel Then Exit For
        Next

    End Sub






    Private Sub ProcessPage_FindSymbol()

        ' IOLIST_READ

        Dim lStat As Boolean

        Dim Ldata As clsCITECTData
        Ldata = New clsCITECTData

        Ldata.ParsePage()

        Dim i As Long
        Dim j As Long
        Dim LResult As DialogResult

        LResult = vbOK

        Dim a As Integer
        Dim b As Integer
        Dim c As Integer


        For i = 1 To Ldata.ObjectCount
            Dim lObj As clsCITECTObject
            lObj = Ldata.GetObject(i)


            Debug.Print(lObj.GetCitectObjectType)
            If LCase(lObj.GetCitectObjectType) = "symbol" Then
                Debug.Print("Symbol = " + lObj.SymbolName + ", " + lObj.LibraryName)
            End If

            If LCase(lObj.GetCitectObjectType) = "set" Then
                lObj.SelectObj()
                MsgBox("press OK to GO On")
            End If



        Next

    End Sub

    Private Sub cmdAdjustSFINumbers_Click()

        If Not CurrentProject() = modCustom.CustomProjectName Then
            MsgBox("Sorry, but the current project in citect(" + CurrentProject() + ") does not match the custom project in 'modCustom' (" + modCustom.CustomProjectName + ")")
            Exit Sub
        End If

        If ConfirmDangerousOperation("This will modify ALL sfi numbers on current citect page. Are you sure?") = False Then Exit Sub




        Dim lStat As Boolean

        Dim Ldata As clsCITECTData
        Ldata = New clsCITECTData

        Ldata.ParsePage()


        Dim i As Long
        Dim j As Long
        Dim LResult As DialogResult

        LResult = vbOK

        For i = 1 To Ldata.GenieCount
            Dim lGen As clsCITECTGenie
            Dim Lprop As clsCITECTProperty

            lGen = Ldata.GetGenie(i)

            Dim lNewVal As String


            Lprop = lGen.GetPropertyByName("SFI_NUMBER")

            lNewVal = modCustom.CustomAdjustSFINumber(lGen.ObjectName, lGen.Get_SFI_NUMBER)
            If Not lNewVal = "" Then
                lGen.UpdateProperty("SFI_NUMBER", lNewVal)
                Debug.Print(lNewVal)
            Else
                Debug.Print("***SKIPPED :" + lGen.ObjectName)
            End If

        Next

    End Sub

    Private Sub cmdClacAddress_Click()
        modConfiguration.LoadConfiguration()
        modTagGeneration.CalcAddress()
    End Sub

    Private Sub cmdCalcAddressed_Click()
        modConfiguration.LoadConfiguration()
        modTagGeneration.CalcAddress()
    End Sub

    Private Sub cmdCheckAlarms_Click()
        modIOList.CompareIOListToAlarmConfiguration()
    End Sub

    Private Sub cmdCheckForDuplicateTags_Click()
        modIOList.CheckForDuplicateTags()
    End Sub

    Private Sub cmdClearAllAlarms_Click()

        ' IOLIST_READ

        If ConfirmDangerousOperation("This will clear ALL alarmnumbers on current page. Are you SURE?") = False Then Exit Sub

        Dim lStat As Boolean

        Dim Ldata As clsCITECTData
        Ldata = New clsCITECTData

        Ldata.ParsePage()


        Dim i As Long
        Dim j As Long

        For i = 1 To Ldata.GenieCount
            Dim lGen As clsCITECTGenie
            lGen = Ldata.GetGenie(i)


            Dim lSys1 As String
            Dim lSys2 As String

            If lGen Is Nothing Then
                MsgBox("what!!!")
            ElseIf lGen.GetPropertyValueByName("AlarmNo") <> "" Then
                lGen.UpdateProperty("AlarmNo", "")
            End If


        Next

    End Sub
    Private Sub cmdGetAllAlarms_Click()


    End Sub


    Private Sub cmdClearAllAlarms_718_Bilge()

        Dim lStat As Boolean

        Dim Ldata As clsCITECTData
        Ldata = New clsCITECTData

        Ldata.ParsePage()


        Dim i As Long
        Dim j As Long

        For i = 1 To Ldata.GenieCount
            Dim lObj As clsCITECTObject
            lObj = Ldata.GetObject(i)

            If lObj Is Nothing Then
                MessageBox.Show("what!!!")
            ElseIf LCase(lObj.GetCitectObjectType) = "squarev3" Then
                lObj.SetFillcolorExpression("")
            End If


        Next

    End Sub

    Public Sub cmdRemoveSpesificProperty(objectProperty As String)

        If ConfirmDangerousOperation("This will remove ALL tooltips current page. Are you SURE?") = False Then Exit Sub
        Dim Ldata As clsCITECTData
        Ldata = New clsCITECTData
        Dim lGen As clsCITECTGenie
        Dim GraphicsBuilder As GraphicsBuilderClass = New GraphicsBuilderClass()
        Ldata.ParsePage()


        For i = 1 To Ldata.GenieCount
            lGen = Ldata.GetGenie(i)

            'If Not lGen.GetPropertyByName("TOOLTIP") Is Nothing Then
            '("ToolTip") Then 'And lGen.GetPropertyValueByName = ("") Then

            If lGen.GetPropertyValueByName(objectProperty) <> "" Then
                lGen.UpdateProperty(objectProperty, "")
            End If

            'Dim lObj As clsCITECTObject
            'lObj = Ldata.GetObject(i)
            'lObj.SelectObj()

            'GraphicsBuilder.sele
            'End If

        Next
        MessageBox.Show("All tooltips removed")

    End Sub

    Public Sub cmdClearLineFillColorExpression_Click()

        If ConfirmDangerousOperation("This will remove ALL linecoloring expressions from current page. Are you SURE?") = False Then Exit Sub
        Dim lStat As Boolean
        Dim Ldata As clsCITECTData = New clsCITECTData
        Dim objectCount As Integer
        Ldata.ParsePage()
        Dim i As Long
        'Dim j As Long

        'MessageBox.Show("Total objects: " + Ldata.ObjectCount.ToString)
        objectCount = Ldata.ObjectCount - Ldata.GenieCount
        For i = 1 To Ldata.ObjectCount
            Dim lObj As clsCITECTObject
            lObj = Ldata.GetObject(i)

            'MessageBox.Show(lObj.GetCitectObjectType)

            If LCase(lObj.GetCitectObjectType) = "line" Then
                lObj.SelectObj()
                lObj.SetFillcolorExpression("")
                'MessageBox.Show("Removed Line object fill.")
            End If

            If LCase(lObj.GetCitectObjectType) = "pipe" Then
                lObj.SelectObj()
                lObj.SetFillcolorExpression("")
                'MessageBox.Show("Removed Pipe object fill.")
            End If

            If lObj.GetCitectObjectType = "CircleV2" Then
                lObj.SelectObj()
                lObj.SetFillcolorExpression("")
                'MessageBox.Show("Removed CircleV2 object fill.")
            End If

            'MessageBox.Show("Nothing done.")

            QuickEditCitect.ProgressBar1.Maximum = objectCount
            QuickEditCitect.ProgressBar1.Minimum = 0
            QuickEditCitect.ProgressBar1.Step = 1
            QuickEditCitect.ProgressBar1.PerformStep()

        Next
        MessageBox.Show("Removed all fill properties on current page.")
        QuickEditCitect.ProgressBar1.Value = 0
    End Sub

    Private Sub cmdCompareIOLists_Click()
        Dim ListB As String
        Dim ListA As String
        ListB = "C:\WINDOWS\Profiles\svein.ingebrigtsen\Desktop\SCRIPTS\741-792-008 IO-List- As Built.xls"
        ListA = "C:\WINDOWS\Profiles\svein.ingebrigtsen\Desktop\SCRIPTS\741-792-008 IO-List-B Preliminary.xls"
        modIOList.CompareIOLists(ListA, ListB)
    End Sub


    Public Sub cmdFindUntaggedGenies_Click()
        If MsgBox("Are you sure you want to start finding untagged objects on the current page?", vbYesNo + vbExclamation) _
            = MsgBoxResult.Cancel Then Exit Sub

        Dim lStat As Boolean

        Dim Ldata As clsCITECTData
        Ldata = New clsCITECTData
        Dim untaggedLines As clsCITECTData
        untaggedLines = New clsCITECTData
        Dim untaggedGenies As Collection
        untaggedGenies = New Collection

        Dim lGen As clsCITECTGenie
        Dim Lprop As clsCITECTProperty

        Ldata.ParsePage()

        Dim i As Long
        Dim j As Long
        Dim LResult As DialogResult

        LResult = vbOK

        For i = 1 To Ldata.GenieCount
            Try
                lGen = Ldata.GetGenie(i)

                If lGen.IsTagged = False Then
                    untaggedGenies.Add(lGen)
                End If

                'If lGen.GetPropertyByName("AlarmNo") Is Nothing _
                '    And lGen.GetPropertyByName("Alarm No") Is Nothing _
                '    And lGen.GetPropertyByName("Alarm No:") Is Nothing _
                '    And lGen.GetPropertyByName("alarm no:") Is Nothing _
                '    And lGen.GetPropertyByName("TagName:") Is Nothing _
                '    And lGen.GetPropertyByName("tagName") Is Nothing _
                '    And lGen.GetPropertyByName("TOOLTIP") Is Nothing _
                '    And lGen.GetPropertyByName("FileName") Is Nothing _
                '    And lGen.GetPropertyByName("DrawNo") Is Nothing _
                '  And lGen.GetPropertyByName("DrawName") Is Nothing Then


                '    untaggedGenies.Add(lGen)

                'End If

            Catch ex As Exception
                MessageBox.Show("An error occured at: " + i.ToString + "                       Exception: " + ex.ToString)
            End Try


        Next

        If untaggedGenies.Count = 0 Then
            MessageBox.Show("There was no untagged genies")
            Exit Sub


        ElseIf MsgBox("Number of objects without a tag: " + untaggedGenies.Count.ToString +
                  " Do you want to go trough them now?", vbYesNo + vbExclamation) _
            = MsgBoxResult.No Then
            Exit Sub
        End If

        For i = 1 To untaggedGenies.Count
            lGen = untaggedGenies(i)
            lGen.SelectObject()
            If MsgBox("Selected object is missing tags. Press OK to continue. ", vbOKCancel) = MsgBoxResult.Cancel Then
                QuickEditCitect.ProgressBar1.Value = 0
                QuickEditCitect.Label1.Visible = False
                Exit Sub
            End If
            QuickEditCitect.ProgressBar1.Maximum = untaggedGenies.Count
            QuickEditCitect.ProgressBar1.Minimum = 0
            QuickEditCitect.ProgressBar1.Step = 1
            QuickEditCitect.ProgressBar1.PerformStep()
            QuickEditCitect.Label1.Visible = True
            QuickEditCitect.Label1.Text = i.ToString + "/" + untaggedGenies.Count.ToString

        Next


        If MsgBox("Finnished. Do you want to recheck for any genies missing tags? ", vbYesNo + vbExclamation) = MsgBoxResult.No Then
            QuickEditCitect.ProgressBar1.Value = 0
            QuickEditCitect.Label1.Visible = False
            Exit Sub
        Else
            QuickEditCitect.ProgressBar1.Value = 0
            QuickEditCitect.Label1.Visible = False
            cmdFindUntaggedGenies_Click()
        End If
    End Sub

    Private Sub cmdGenerateAkerDBF_Click()
        modConfiguration.LoadConfiguration()
        modTagGeneration.FillAker_dbf()
    End Sub

    Private Sub cmdGenerateAlarmAddresses_Click()
        modConfiguration.LoadConfiguration()
        modTagGeneration.AlarmAddress()
    End Sub

    Private Sub cmdGenerateAlarms_Click()
        modConfiguration.LoadConfiguration()
        modTagGeneration.Alarm()
    End Sub

    Private Sub cmdGenerateAlarmCSV_Click()
        modConfiguration.LoadConfiguration()
        modTagGeneration.AlarmCsvGen()
    End Sub

    Private Sub cmdGenerateIsagraf_Click()
        modConfiguration.LoadConfiguration()
        modTagGeneration.IsaGraf()
    End Sub

    Private Sub cmdGetMaxAlarmNo_Click()
        modConfiguration.LoadConfiguration()
        utilities.GetNextAlarmNo()
    End Sub




    Private Sub cmdSetMissingAlarms_Click()

        If ConfirmDangerousOperation("This will set alarmnumber to '3999' for genies not yet assigned an alarm number. Are you SURE?") = False Then Exit Sub

        Dim lStat As Boolean

        Dim Ldata As clsCITECTData
        Ldata = New clsCITECTData

        Ldata.ParsePage()


        Dim i As Long
        Dim j As Long

        For i = 1 To Ldata.GenieCount
            Dim lGen As clsCITECTGenie
            lGen = Ldata.GetGenie(i)

            If Not lGen.GetPropertyByName("AlarmNo") Is Nothing And lGen.GetPropertyValueByName("AlarmNo") = "" Then
                lGen.UpdateProperty("AlarmNo", "3999")
                Debug.Print("Missing alarm: " + lGen.ObjectName + ", SFI:" + lGen.GetPropertyValueByName("SFI_NUMBER") + ", MODBUS:" + lGen.GetPropertyValueByName("MODBUS_ADDRESS"))
            End If
        Next

    End Sub

    Private Sub cmdTank_Click()
        modConfiguration.LoadConfiguration()
        modTagGeneration.tank()
    End Sub

    Private Sub cmdTest_Click()
        'Dim lLibs As clsCitectLibraries
        'lLibs = New clsCitectLibraries

        'lLibs.ParseLibraries

    End Sub

    Private Sub cmdtransferDataToIOList_Click()
        If ConfirmDangerousOperation("This may mess up your IO list. Are you sure?") = False Then Exit Sub

        modIOList.TransferSettings("F:\Installasjon\Prosjekt\STX-ROB\740\Automation\IAS\IO List\SHI - UPDATE\OLD\COPY_IO-List-B740-Rev.A.xls", "BK")

    End Sub

    Private Sub cmdValve_Click()
        modConfiguration.LoadConfiguration()
        modTagGeneration.valve()
    End Sub


    Private Sub CommandButton1_Click()
        ' IOLIST_READ

        If ConfirmDangerousOperation("This will clear ALL alarmnumbers on current page. Are you SURE?") = False Then Exit Sub

        Dim lStat As Boolean
        Dim sAlmNo As String

        Dim Ldata As clsCITECTData
        Ldata = New clsCITECTData

        Ldata.ParsePage()


        Dim i As Long
        Dim j As Long

        For i = 1 To Ldata.GenieCount
            Dim lGen As clsCITECTGenie
            lGen = Ldata.GetGenie(i)


            Dim lSys1 As String
            Dim lSys2 As String

            If lGen Is Nothing Then
                MsgBox("what!!!")
            ElseIf lGen.GetPropertyValueByName("AlarmNo") <> "" Then
                sAlmNo = "ALM" & lGen.GetPropertyValueByName("AlarmNo")
            End If


        Next

    End Sub


End Class
