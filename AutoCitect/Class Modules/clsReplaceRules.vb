Option Explicit On
Imports Microsoft.Office.Interop

Public Class clsReplaceRules

    Private m_Page As clsCITECTData
    Private m_IOList As IO_List
    Private m_Rules As Collection
    Private m_RuleSheet As Excel.Worksheet


    Public Sub LoadRules(piWS As Excel.Worksheet)
        Dim i As Long
        i = 2
        While piWS.Cells(i, "B") <> ""
            LoadRule(piWS, i)
            i = i + 1
        End While
    End Sub

    Private Sub LoadRule(piWS As Excel.Worksheet, piIndex As Long)
        Dim lRule As clsReplaceRule

        m_RuleSheet = piWS


        lRule = Add()
        lRule.SystemName = piWS.Cells(piIndex, "B")
        lRule.Description = piWS.Cells(piIndex, "C")

        lRule.KeyPropertyName = piWS.Cells(piIndex, "I")

        Select Case LCase(lRule.KeyPropertyName)
            Case "sfi_number"
                lRule.KeyPropertyValue = piWS.Cells(piIndex, "A")
            Case "description"
                lRule.KeyPropertyValue = piWS.Cells(piIndex, "C")
            Case "modbus_address"
                lRule.KeyPropertyValue = piWS.Cells(piIndex, "E")
            Case "tag"
                lRule.KeyPropertyValue = piWS.Cells(piIndex, "G")
        End Select

        lRule.AddPropertiesToUpdate(piWS.Cells(piIndex, "J"))
        lRule.ExcelRow = piIndex
        lRule.INTERNALSetRuleContentsStatus( lRule.RuleContentsStatus.REPLACE_RULE_EXISTING)
    End Sub


    Public Sub Execute(piIoList As IO_List, piPage As clsCITECTData)
        Dim lRule As clsReplaceRule
        Dim lFeedback As String
        Dim lMaxrow As Long
        lMaxrow = 0

        For Each lRule In m_Rules
            lRule.Execute(piPage, piIoList, lFeedback)
            If Not lRule.IsCompleted Then
                MsgBox("Error: ---update not completed",
                                 vbOKOnly + vbCritical, "Error")
            End If
            m_RuleSheet.Cells(lRule.ExcelRow, "M") = lFeedback
            If lFeedback = "OKOK" Then
                lFeedback = lFeedback
            End If
            If lRule.ExcelRow >= lMaxrow Then lMaxrow = lRule.ExcelRow + 1
        Next

        'Finally report untagged genies-----
        Dim i As Long
        Dim lGen As clsCITECTGenie
        For i = 1 To piPage.GenieCount
            lGen = piPage.GetGenie(i)
            lFeedback = ""
            If lGen.IsTagged = False Then
                lFeedback = "UNTAGGED GENIE: " + lGen.ObjectName + ":" + lGen.Get_SFI_NUMBER
                m_RuleSheet.Cells(lMaxrow, "M") = lFeedback
                lMaxrow = lMaxrow + 1
            End If
        Next
        '-----------------------------------

    End Sub

    Public Function Add() As clsReplaceRule
        Dim lRetval As clsReplaceRule
        lRetval = New clsReplaceRule
        m_Rules.Add(lRetval)
        Add = lRetval
    End Function

    Public ReadOnly Property RulesCount() As Long
        Get
            RulesCount = m_Rules.Count
        End Get
    End Property

    Public Function GetRule(piIndex As Long) As clsReplaceRule
        On Error Resume Next
        GetRule = m_Rules(piIndex)
    End Function

    Public Function GetSystemNames() As Collection
        Dim lRetval As Collection
        lRetval = New Collection


        Dim lRule As clsReplaceRule
        For Each lRule In m_Rules
            On Error Resume Next
            lRetval.Add(lRule.SystemName, lRule.SystemName)
        Next

        Dim lRetval2 As Collection
        lRetval2 = New Collection
        Dim lSysName As Object
        For Each lSysName In lRetval
            lRetval2.Add(lSysName)
        Next

        GetSystemNames = lRetval2

    End Function


    Private Sub Class_Initialize()
        m_Rules = New Collection
    End Sub

    Private Sub UpdateFromIOList(piIoList As IO_List)
        'DO NOT USE!!!
        Dim lRule As clsReplaceRule

        '    Dim lSystNames As Collection
        '    Set lSystNames = GetSystemNames()

        Dim lNewSignals As Collection
        lNewSignals = New Collection
        'FIST WALK THROUGH ALL EXISTING RULES
        For i = 1 To Me.RulesCount
            Dim lRule2 As clsReplaceRule
            lRule = Me.GetRule(i)

            Dim lSig As IO_List_Signal
            lSig = piIoList.GetSignalBy(lRule.SystemName, lRule.KeyPropertyName, lRule.KeyPropertyValue)
            If lSig Is Nothing Then
                lRule2.INTERNALSetRuleContentsStatus(lRule2.RuleContentsStatus.REPLACE_RULE_DELETED)
            Else
                lRule2.INTERNALSetRuleContentsStatus(lRule2.RuleContentsStatus.REPLACE_RULE_UPDATED)
                lSig.SignalInUse = True
            End If
        Next

        'THEN LOOK FOR NEW SIGNALS
        Dim lSystems As Collection
        lSystems = Me.GetSystemNames
        For i = 1 To lSystems.Count
            Dim lSys As IO_List_System

            lSys = piIoList.GetSystemByName(lSystems(i))
            Dim lSig As IO_List_Signal

            Dim j As Long
            For j = 1 To lSys.SignalCount
                lSig = lSys.GetSignal(j)
                If lSig.SignalInUse = False Then
                    AddruleFromSignal(lSys.SYSTEM_NAME, lSig)
                End If
            Next
        Next

    End Sub



    Private Sub AddruleFromSignal(piSys As String, piSig As IO_List_Signal)
        'DO NOT USE!!!
        Dim lRule As clsReplaceRule
        lRule = Me.Add
        lRule.SystemName = piSys
        lRule.Description = piSig.Description
    End Sub

End Class
