Option Explicit On
Public Class clsAlarms
    Private m_Alarms As Collection

    Private Sub Class_Initialize()
        m_Alarms = New Collection
    End Sub

    Public Sub Initialize(piMaxNoAlarms As Long, piIoList As IO_List)
        m_Alarms = New Collection

        Dim i As Long
        Dim lAlarm As clsAlarm
        Dim lAlmNo As Long

        For i = 1 To piMaxNoAlarms
            lAlarm = New clsAlarm
            lAlarm.AlarmNumber = i
            m_Alarms.Add(lAlarm)
        Next

        Dim lSig As IO_List_Signal
        For i = 1 To piIoList.SignalCount
            lSig = piIoList.getSignalByIndex(i)

            If Left(lSig.IASTagname, 3) = "Alm" Then

                lAlmNo = Right(lSig.IASTagname, 4)

                lAlarm = m_Alarms(lAlmNo)
                lAlarm.Signal = lSig
            End If
        Next

    End Sub

    Public ReadOnly Property AlarmCount() As Long
        Get
            AlarmCount = m_Alarms.Count
        End Get
    End Property

    Public Function GetAlarm(piAlarmNo As Long) As clsAlarm
        On Error Resume Next
        GetAlarm = m_Alarms(piAlarmNo)
    End Function

End Class
