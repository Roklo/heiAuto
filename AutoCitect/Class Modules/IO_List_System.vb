Option Explicit On
Imports Microsoft.Office.Interop

Public Class IO_List_System
    Public SYSTEM_NAME As String

    Private m_Signals As Collection


    Private Sub Class_Initialize()
        m_Signals = New Collection

    End Sub

    Public ReadOnly Property SignalCount() As Long
        Get
            SignalCount = m_Signals.Count
        End Get
    End Property

    Public Function GetSignal(piIndex As Long) As IO_List_Signal
        On Error Resume Next
        GetSignal = m_Signals(piIndex)
    End Function

    Public Function GetSignalBy(PiPropertyName As String, piValue As String) As IO_List_Signal
        Dim lSig As IO_List_Signal
        If piValue = "" Then Exit Function

        For Each lSig In m_Signals
            If lSig.GetPropertByName(PiPropertyName) = piValue Then
                GetSignalBy = lSig
                Exit For
            End If
        Next
    End Function

    Public Function Read(piWS As Excel.Worksheet, ByRef cur_line As Long, piIoList As IO_List)
        If SYSTEM_NAME = "" Then
            SYSTEM_NAME = GetSystemName(piWS, cur_line)
        End If

        Dim lSysName As String
        lSysName = SYSTEM_NAME

        Dim lSig As IO_List_Signal

        While SYSTEM_NAME = lSysName And Not lSysName = ""
            lSig = New IO_List_Signal
            m_Signals.Add(lSig)

            lSig.Read(piWS, cur_line, piIoList)


            cur_line = cur_line + 1
            lSysName = GetSystemName(piWS, cur_line)
        End While

    End Function


    Private Function GetSystemName(piWS As Excel.Worksheet, line As Long)
        GetSystemName = piWS.Cells(line, 6)
    End Function

End Class
