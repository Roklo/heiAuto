Option Explicit On

Public Class clsAlarm
    Public Enum NORMAL_VALUE
        NORMAL_VALUE_OPEN = 1
        NORMAL_VALUE_CLOSED = 0
        NORMAL_VALUE_UNKNOWN = -1
    End Enum


    Public AlarmNumber As String
    Public Signal As IO_List_Signal


    Public ReadOnly Property Name() As String
        Get
            If Not Me.Signal Is Nothing Then
                Name = Me.Signal.SFI_NUMBER.SFI
            End If
            If Name = "" Then
                Name = "Alm" & Format(AlarmNumber, "000#") & "_NAME"
            End If
        End Get
    End Property

    Public ReadOnly Property Description() As String
        Get
            Dim lDesc As String
            If Not Me.Signal Is Nothing Then
                lDesc = Me.Signal.Description
            End If

            If lDesc = "" Then
                lDesc = "Alm" & Format(AlarmNumber, "000#") & "_DESC"
            End If

            Description = lDesc
        End Get
    End Property

    Public ReadOnly Property Unit() As String
        Get
            If Not Me.Signal Is Nothing Then
                Unit = Me.Signal.Unit
            End If
            If Unit = "" Then
                Unit = "_"
            End If
        End Get
    End Property

    Public ReadOnly Property SignalType() As String
        Get
            If Me.Signal Is Nothing Then Exit Property
            SignalType = Signal.SignalType
        End Get
    End Property

    Public ReadOnly Property OnTxt() As String
        Get
            If Me.Signal Is Nothing Then
                OnTxt = "Alarm"
                Exit Property
            End If
            If UCase(Left(Signal.SignalType, 2)) = "DI" Then
                If Signal.NormalValue = "NO" Then
                    OnTxt = "Alarm"
                ElseIf Signal.NormalValue = "NC" Then
                    OnTxt = "Normal"
                End If
            Else
                OnTxt = "Alarm"
            End If
        End Get
    End Property

    Public ReadOnly Property OffTxt() As String
        Get
            If Me.Signal Is Nothing Then
                OffTxt = "Normal"
                Exit Property
            End If

            If UCase(Left(Signal.SignalType, 2)) = "DI" Then
                If Signal.NormalValue = "NO" Then
                    OffTxt = "Normal"
                ElseIf Signal.NormalValue = "NC" Then
                    OffTxt = "Alarm"
                End If
            Else
                OffTxt = "Normal"
            End If
        End Get
    End Property

    Public ReadOnly Property NormalValue() As NORMAL_VALUE

        Get
            Dim NORMAL_VALUE_CLOSED As String = ""
            Dim NORMAL_VALUE_OPEN As String = ""
            NormalValue = NORMAL_VALUE_CLOSED
            If Me.Signal Is Nothing Then
                Exit Property
            End If

            If UCase(Me.Signal.NormalValue) = "NO" Then
                NormalValue = NORMAL_VALUE_OPEN
            End If
        End Get
    End Property

    Private Sub Class_Initialize()
        Signal = New IO_List_Signal
    End Sub

End Class
