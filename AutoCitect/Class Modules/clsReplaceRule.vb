Option Explicit On

Public Class clsReplaceRule
    Public Enum RuleContentsStatus
        REPLACE_RULE_UNKNOWN
        REPLACE_RULE_EXISTING
        REPLACE_RULE_NEW
        REPLACE_RULE_DELETED
        REPLACE_RULE_UPDATED
    End Enum


    Private m_PropertiesToUpdate As Collection

    Public Description As String
    Public SystemName As String

    Public KeyPropertyName As String
    Public KeyPropertyValue As String

    Public ExcelRow As Long

    Private m_Completed As Boolean

    Private m_RuleContentsStatus As RuleContentsStatus

    Public ReadOnly Property GetRuleContentsStatus() As RuleContentsStatus
        Get
            GetRuleContentsStatus = m_RuleContentsStatus
        End Get
    End Property

    Public Function INTERNALSetRuleContentsStatus(piStatus As RuleContentsStatus)
        ' Set
        m_RuleContentsStatus = piStatus
        'End Set

    End Function


    Public ReadOnly Property IsCompleted() As Boolean
        Get
            IsCompleted = m_Completed
        End Get
    End Property

    Private Sub Class_Initialize()
        m_PropertiesToUpdate = New Collection
        m_RuleContentsStatus = RuleContentsStatus.REPLACE_RULE_UNKNOWN
        ExcelRow = 0
    End Sub

    Public ReadOnly Property PropertiesToUpdateCount() As Long
        Get
            PropertiesToUpdateCount = m_PropertiesToUpdate.Count
        End Get
    End Property


    Public Function GetPropertyToUpdate(piIndex As Long) As Long
        On Error Resume Next
        GetPropertyToUpdate = m_PropertiesToUpdate(piIndex)
    End Function

    Public Function AddPropertyToUpdate(piName As String)
        m_PropertiesToUpdate.Add(piName)
    End Function

    Public Function AddPropertiesToUpdate(piCommaseparatedNames As String)
        Dim lProps As Collection
        lProps = SplitStringByComma(piCommaseparatedNames)

        Dim Lprop As Object
        For Each Lprop In lProps
            Dim lString As String
            lString = Lprop
            AddPropertyToUpdate(lString)
        Next
    End Function


    Public Sub Execute(piPage As clsCITECTData, piIoList As IO_List, ByRef poResultStatus As String)

        poResultStatus = ""
        If piPage Is Nothing Or piIoList Is Nothing Then Exit Sub

        m_Completed = True



        Dim lSignal As IO_List_Signal
        lSignal = piIoList.GetSignalBy(SystemName, KeyPropertyName, KeyPropertyValue)

        If "30091" = KeyPropertyValue Then
            KeyPropertyValue = KeyPropertyValue
        End If



        If lSignal Is Nothing Then
            MsgBox("clsReplaceRule::Execute(), did not find signal:" + SystemName + ":" + KeyPropertyName + ":" + KeyPropertyValue,
                                 vbOKOnly + vbCritical, "Error")
            poResultStatus = "SIGNAL_NOT_FOUND"
            m_Completed = False
            Exit Sub
        End If

        Dim lGenie As clsCITECTGenie
        Dim lGenies As Collection

        lGenies = piPage.FindGeniesBy(KeyPropertyName, KeyPropertyValue)
        If lGenies Is Nothing Then
            MsgBox("clsReplaceRule::Execute(), did not find signal:" + SystemName + ":" + KeyPropertyName + ":" + KeyPropertyValue,
                                 vbOKOnly + vbCritical, "Error")
            poResultStatus = "GENIE_NOT_FOUND"
            m_Completed = False
            Exit Sub
        ElseIf lGenies.Count = 0 Then
            Debug.Print("clsReplaceRule::Execute(), did not find genie: " + SystemName + ":" + KeyPropertyName + ":" + KeyPropertyValue)
            poResultStatus = "GENIE_NOT_FOUND"
            m_Completed = False
            Exit Sub
        End If

        Dim lVar As Object
        Dim lPropertyName As String
        Dim lGenieCompleted As Boolean

        For Each lGenie In lGenies
            lGenieCompleted = True


            For Each lVar In m_PropertiesToUpdate
                lPropertyName = lVar
                Dim lPropertyValue As String
                lPropertyValue = lSignal.GetPropertByName(lPropertyName)

                If lPropertyValue = "" Then
                    Debug.Print("clsReplaceRule::Execute(), did not find propertyvalue in signal: " + SystemName + ":" + KeyPropertyName + "(" + KeyPropertyValue + ") Missing property=" + lPropertyName)
                    If Len(poResultStatus) > 0 Then poResultStatus = poResultStatus + ","
                    poResultStatus = poResultStatus + "SIGNAL_Property(" + lPropertyName + ")NOT FOUND"
                    lGenieCompleted = False
                    m_Completed = False
                End If



                If lGenie.GetPropertyByName(lPropertyName) Is Nothing Then
                    m_Completed = False
                    lGenieCompleted = False

                    If Len(poResultStatus) > 0 Then poResultStatus = poResultStatus + ","
                    poResultStatus = poResultStatus + "GENIE_Property(" + lPropertyName + ")NOT FOUND"

                End If

                lGenie.UpdateProperty(lPropertyName, lPropertyValue)
            Next 'property

            If lGenieCompleted Then
                '*****APPEND A STAR TO THE DESCRIPTION TO SHOW IT HAS BEED UPDATED*****
                Dim lDesc As String
                lDesc = "*" + lGenie.GetPropertyValueByName("ALARM DESCRIPTION")
                If (Len(lDesc) > 1) Then
                    lGenie.UpdateProperty("ALARM DESCRIPTION", lDesc)
                End If

                lDesc = "*" + lGenie.GetPropertyValueByName("Description")
                If (Len(lDesc) > 1) Then
                    lGenie.UpdateProperty("Description", lDesc)
                End If

                lGenie.MarkAsTagged(lGenieCompleted)

                poResultStatus = poResultStatus + "OK"
                '-----------------------------------------------------------------------
            End If
        Next 'genie
    End Sub

End Class
