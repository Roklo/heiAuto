Option Explicit On
Imports GraphicsBuilder

Public Class clsCoordinates
    Dim GraphicsBuilder As GraphicsBuilderClass = New GraphicsBuilderClass()
    Private m_Coordinates As Collection = New Collection
    Private m_ObjType As String

    Private Sub Class_Initialize()
        m_Coordinates = New Collection
    End Sub


    Public ReadOnly Property count() As Long
        Get
            count = m_Coordinates.Count
        End Get
    End Property

    Public Function GetCoordinate(piIndex As Long) As clsCoordinate
        On Error Resume Next
        GetCoordinate = m_Coordinates(piIndex)
    End Function

    Public Function Add(piX As Integer, piY As Integer) As clsCoordinate
        Dim LC As clsCoordinate

        LC = New clsCoordinate
        m_Coordinates.Add(LC)

        LC.X = piX
        LC.Y = piY

        Add = LC
    End Function


    Public Sub ParseObject(piObjType As String)
        m_ObjType = piObjType

        Dim lContinue As Boolean
        lContinue = CitectGetCoordinate(True)

        While lContinue
            lContinue = CitectGetCoordinate(False)
        End While
    End Sub

    Private Function CitectGetCoordinate(first_time As Boolean) As Boolean
        On Error GoTo ERR_HANDLER
        Dim lX As Integer
        Dim lY As Integer

        If m_ObjType = "Line" Or m_ObjType = "CircleV2" Then
            Dim X2 As Integer
            Dim Y2 As Integer

            GraphicsBuilder.AttributeBaseCoordinates(lX, lY, X2, Y2)

            Me.Add(lX, lY)
            Me.Add(X2, Y2)

        ElseIf m_ObjType = "Polygon" Or m_ObjType = "Pipe" Or m_ObjType = "Free Hand Line" Then
            ' Get coordinates of a free hand line, polygon or pipe.
            If first_time Then
                GraphicsBuilder.AttributeNodeCoordinatesFirst(lX, lY)
            Else
                GraphicsBuilder.AttributeNodeCoordinatesNext(lX, lY)
            End If
            Me.Add(lX, lY)
            CitectGetCoordinate = True
        End If

        Exit Function
ERR_HANDLER:
    End Function


End Class
