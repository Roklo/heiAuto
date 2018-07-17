Option Explicit On
Public Class clsCitectPropertyNames
    Private m_Props As Collection
    Private m_PropsByIndex As Collection

    Private Sub Class_Initialize()
        m_Props = New Collection
        m_PropsByIndex = New Collection
    End Sub


    Public Sub Add(PiProp As clsCITECTProperty)
        Try
            If PiProp Is Nothing Then Exit Sub
            If PiProp.Name = "" Then Exit Sub

            m_Props.Add(PiProp, PiProp.Name)

            m_PropsByIndex.Add(PiProp)
            Exit Sub
        Catch ex As Exception
            'Do nothing, already defined
        End Try
    End Sub

    Public ReadOnly Property PropertyCount() As Long
        Get
            PropertyCount = m_PropsByIndex.Count
        End Get
    End Property


    Public Function GetPropertyName(piIndex As Long) As clsCITECTProperty
        On Error Resume Next
        GetPropertyName = m_PropsByIndex(piIndex)
    End Function


    Public Function GetPropertyNameByName(piName As String) As clsCITECTProperty
        On Error Resume Next
        GetPropertyNameByName = m_Props(piName)
    End Function
End Class
