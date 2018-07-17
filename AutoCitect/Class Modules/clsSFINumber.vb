Option Explicit On
Public Class clsSFINumber

    Private m_First As String
    Private m_Second As String
    Private m_Third As String
    Private m_Tail As String


    Public Property First() As String
        Get
            First = m_First
        End Get
        Set(piFirst As String)
            m_First = piFirst
        End Set


    End Property

    Public Property Second() As String
        Get
            Second = m_Second
        End Get
        Set(piSecond As String)
            m_Second = piSecond
        End Set


    End Property

    Public Property Third() As String
        Get
            Third = m_Third
        End Get
        Set(piThird As String)
            m_Third = piThird
        End Set
    End Property

    Public Property Tail() As String
        Get
            Tail = m_Tail
        End Get
        Set(piTail As String)
            m_Tail = piTail
        End Set
    End Property


    Public Sub SetSFI(piValue As String)

        Dim lStr As String
        Dim lX As Long

        First = ""
        Second = ""
        Third = ""
        Tail = ""

        If Len(piValue) = 0 Then
            Exit Sub
        End If


        Dim lValues As Collection
        lValues = SplitString(piValue)

        On Error GoTo ERR_HANDLER

        First = lValues(1)
        Second = lValues(2)
        Third = lValues(3)
        Tail = lValues(4)


        Exit Sub
ERR_HANDLER:
    End Sub

    Private Function SplitString(piVal As String) As Collection
        Dim result As Object

        result = Split(piVal, ".")
        If UBound(result) < 2 Then
            result = Split(piVal, "_")
        End If

        Dim lRetval As Collection
        lRetval = New Collection

        Dim i As Long
        For i = 0 To UBound(result)
            lRetval.Add(result(i))
        Next

        SplitString = lRetval



    End Function

    Private Function LocateSeparator(piValue As String) As Long
        Dim lX As Long
        lX = InStr(1, piValue, ".")
        If lX < 0 Then
            lX = InStr(1, piValue, "_")
        End If
        LocateSeparator = lX
    End Function


    Public Function IsEqual(piSFI As clsSFINumber) As Boolean
        Dim lRetval As Boolean
        lRetval = True

        If Not piSFI.First = First Then lRetval = False
        If Not piSFI.Second = Second() Then lRetval = False
        If Not piSFI.Third = Third Then lRetval = False
        If Not piSFI.Tail = Tail Then lRetval = False

        IsEqual = lRetval
    End Function


    Public Function IsSameUnit(piSFI As clsSFINumber) As Boolean
        Dim lRetval As Boolean
        lRetval = True

        If Not piSFI.First = First Then lRetval = False
        If Not piSFI.Second = Second() Then lRetval = False

        If Len(Third) > 0 And Len(piSFI.Third) > 0 Then
            If Not Mid(Third, 1, 1) = Mid(piSFI.Third, 1, 1) Then lRetval = False
        ElseIf Len(Third) <> Len(piSFI.Third) Then
            lRetval = False
        End If

        IsSameUnit = lRetval
    End Function

    Public ReadOnly Property SFI() As String
        Get
            If Len(First) > 0 Then SFI = First
            If Len(Second) > 0 Then SFI = SFI + "_" + Second
            If Len(Third) > 0 Then SFI = SFI + "_" + Third
            If Len(Tail) > 0 Then SFI = SFI + "_" + Tail
        End Get
    End Property

End Class
