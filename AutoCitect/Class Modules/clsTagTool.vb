Option Explicit On

Public Class clsTagTool

    Private m_IOList As IO_List

    Public Sub Initialize(piIoList As String)

        m_IOList = New IO_List
        m_IOList.ReadIOList(piIoList) ' "F:\Installasjon\Prosjekt\Aker Aukra\B 706 Aukra\Automation\AYAS\7. Project Tag Lists\IAS\SHI___DO_NOT_USE__THIS_STX NOA 706 INSTRUMENTLISTE REV 2 AYAS REVISION.xls", 1

    End Sub


    Public Sub ProcessPage()

        Dim lStat As Boolean

        Dim Ldata As clsCITECTData
        Ldata = New clsCITECTData

        Ldata.ParsePage()

        Dim i As Long

        Dim lGen As clsCITECTGenie

        For i = 1 To Ldata.GenieCount

            lGen = Ldata.GetGenie(i)
            Dim lSys2 As String = "" 'Addded 06.07.18 by Robin to get rid of comile error

            thr_alarm(lGen, lSys2)
        Next

    End Sub



    Private Sub thr_alarm(piGenie As clsCITECTGenie, piSys As String)

        If piGenie.GetAnimationNumber2 = 798 Then
            Dim X As Long
            X = 1

        End If


        Dim lSignal As IO_List_Signal
        lSignal = FIndSignal(piGenie, piSys, "MODBUS_ADDRESS")

        If lSignal Is Nothing Then
            lSignal = FIndSignal(piGenie, piSys, "SFI_NUMBER")
        End If

        If lSignal Is Nothing Then
            piGenie.SelectObject()
            ' MsgBox "thr_alarm().. did not find signal for"
            Exit Sub
        End If

        'If piGenie.GetPropertyValueByName("ALARM DESCRIPTION") = "" Then
        '    piGenie.UpdateProperty "ALARM DESCRIPTION", lSignal.DESCRIPTION
        'End If

        'If piGenie.GetPropertyValueByName("MODBUS_ADDRESS") = "" Then
        piGenie.UpdateProperty("MODBUS_ADDRESS", lSignal.MODBUS_ADDRESS)
        'End If

        'If piGenie.GetPropertyValueByName("SFI_NUMBER") = "" Then
        piGenie.UpdateProperty("SFI_NUMBER", lSignal.SFI_NUMBER.ToString)
        'End If

        'If piGenie.GetPropertyValueByName("AlarmNo") = "" Then
        piGenie.UpdateProperty("AlarmNo", lSignal.AlarmNo)
        'End If


    End Sub

    Private Sub SetNonexistentAlarmsToZero()

        'IOLIST_READ   : What is this ? Robin


        Dim lStat As Boolean

        Dim Ldata As clsCITECTData
        Ldata = New clsCITECTData

        Ldata.ParsePage()


        Dim i As Long
        Dim j As Long

        For i = 1 To Ldata.GenieCount
            Dim lGen As clsCITECTGenie
            lGen = Ldata.GetGenie(i)
            Dim lDoit As Boolean
            lDoit = False
            Select Case lGen.ObjectName
                Case "bargraph_general"
                    lDoit = True
                Case "thr_alarm"
                    lDoit = True
                Case "alm_analogind_int"
                    lDoit = True
                Case "exch_bargraph_me"
                    lDoit = True
            End Select


            If lDoit And lGen.GetPropertyValueByName("AlarmNo") = "" Then
                lGen.UpdateProperty("AlarmNo", "0000")
            End If
        Next

    End Sub




    Private Sub alm_analogind_int(piGenie As clsCITECTGenie, piSys As String)
        thr_alarm(piGenie, piSys)
    End Sub

    Private Function FIndSignal(piGenie As clsCITECTGenie, piSys As String, piProperty As String) As IO_List_Signal

        Dim lValue As String
        Dim lSFI As String
        Dim lSystem As String

        lValue = piGenie.GetPropertyValueByName(piProperty)
        If lValue = "" Then Exit Function
        '    lSFI = piGenie.GetPropertyValueByName("SFI_NUMBER")

        Dim lIO_Signal As IO_List_Signal
        Dim lIO_Sys As IO_List_System

        If piSys = "" Then

        End If

        lIO_Sys = m_IOList.GetSystemByName(piSys)

        lIO_Signal = lIO_Sys.GetSignalBy(piProperty, lValue)
        FIndSignal = lIO_Signal
    End Function



End Class
