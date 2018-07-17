Option Explicit On
Imports GraphicsBuilder

Public Class clsCITECTProperty
    Private m_Name As String
    Private m_Value As String

    Public ExcelColumn As Long


    Public Property Name() As String
        Get
            Name = m_Name
        End Get

        Set(piVal As String)
            m_Name = piVal
        End Set
    End Property



    Public Property Value() As String
        Get
            Value = m_Value
        End Get
        Set(piVal As String)
            m_Value = piVal
        End Set
    End Property

    Private Sub Class_Initialize()

    End Sub

End Class
