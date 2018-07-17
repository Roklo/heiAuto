Option Explicit On
Imports Test1
Imports Microsoft.VisualBasic
Imports GraphicsBuilder

Public Class clsCITECTObject

    Dim GraphicsBuilder As GraphicsBuilderClass = New GraphicsBuilderClass()

    Private m_AnimationNumber As Integer
    Private m_CitectObjectType As String
    Private m_TextValue As String
    Private m_Description As String
    Private m_Tooltip As String
    Private m_FILL_ON_ColorExpression As String
    Private m_FillColorType As Integer

    Private m_LibraryName As String
    Private m_SymbolName As String
    Private m_ObjectType As Integer

    Private m_Color As Integer
    Private m_ColorLimit As Integer
    Private m_ColorOperator As Integer

    Private m_Coordinates As clsCoordinates


    Private Sub Clear()
        m_AnimationNumber = -1
        m_CitectObjectType = ""
        m_TextValue = ""
        m_Description = ""
        m_Tooltip = ""
        m_FILL_ON_ColorExpression = ""
        m_FillColorType = 0
        m_Coordinates = New clsCoordinates
    End Sub

    Public ReadOnly Property TextValue() As String
        Get
            TextValue = m_TextValue
        End Get
    End Property

    Public ReadOnly Property Description() As String
        Get
            Description = m_Description
        End Get
    End Property



    Public Sub ParseObject()

        Try
            Clear()
            m_AnimationNumber = GetAnimationNumber
            m_CitectObjectType = GetCitectObjectType

            GraphicsBuilder.PropertiesAccessGeneralGet(m_Description, m_Tooltip, 0, 0, "")
            If m_CitectObjectType = "Text" Then
                m_TextValue = GraphicsBuilder.AttributeText
                m_TextValue = m_TextValue
            End If

            Dim lType As Integer
            Dim Lexpression As String


            If m_CitectObjectType = "Line" Then
                GraphicsBuilder.PropertiesFillColourGet(m_FillColorType, m_FILL_ON_ColorExpression, "", "", "", "", False, 0, 0)
                m_Color = GraphicsBuilder.AttributeLineColour

            End If

            If m_CitectObjectType = "Pipe" Then
                GraphicsBuilder.PropertiesFillColourGet(m_FillColorType, m_FILL_ON_ColorExpression, "", "", "", "", False, 0, 0)
                m_Color = GraphicsBuilder.AttributeLineColour
            End If

            If m_CitectObjectType = "CircleV2" Then 'Ellipse
                GraphicsBuilder.PropertiesFillColourGet(m_FillColorType, m_FILL_ON_ColorExpression, "", "", "", "", False, 0, 0)
                m_Color = GraphicsBuilder.AttributeLineColour
            End If

            If LCase(m_CitectObjectType) = "symbol" Then
                'MessageBox.Show(CurrentProject.ToString)
                'MessageBox.Show(CurrentProject() + " " + m_LibraryName + " " + m_SymbolName + " " + GetCitectObjectType)

                'GraphicsBuilder.LibraryObjectName(CurrentProject, m_LibraryName, m_SymbolName, m_ObjectType)

            End If

            If LCase(m_CitectObjectType) = "squarev3" Then
                GraphicsBuilder.PropertiesFillColourGet(m_FillColorType, m_FILL_ON_ColorExpression, "", "", "", "", False, 0, 0)
                m_Color = GraphicsBuilder.AttributeLineColour
            End If


            GetCoordinates()
            Exit Sub
        Catch ex As Exception
            Dim testmsg As Integer
            testmsg = MsgBox("Error: Could not parse object. " + Err.Description,
                                 vbOKOnly + vbCritical, "Error")
        End Try
    End Sub

    Private Sub GetCoordinates()
        m_Coordinates = New clsCoordinates
        m_Coordinates.ParseObject(m_CitectObjectType)
    End Sub


    Public ReadOnly Property GetCitectObjectType() As String
        Get
            GetCitectObjectType = ""
            On Error GoTo ERROR_HANDLER
            If m_CitectObjectType = "" Then
                m_CitectObjectType = GraphicsBuilder.AttributeClass
            End If
ERROR_HANDLER:
            GetCitectObjectType = m_CitectObjectType
            'MessageBox.Show("Error ved å hente AttributeClass")
        End Get
    End Property


    Public ReadOnly Property GetAnimationNumber() As Integer
        Get
            Try
                If m_AnimationNumber <= 0 Then
                    m_AnimationNumber = GraphicsBuilder.AttributeAN
                End If
                GetAnimationNumber = m_AnimationNumber
                Exit Property
            Catch ex As Exception
                Clear()
                GetAnimationNumber = m_AnimationNumber
            End Try
        End Get
    End Property

    Public ReadOnly Property GetFillcolorExpression() As String
        Get
            GetFillcolorExpression = m_FILL_ON_ColorExpression
        End Get
    End Property

    Public ReadOnly Property GetLineColor() As Integer
        Get
            GetLineColor = m_Color
        End Get
    End Property


    Public Function SetFillcolorExpression(piValue As String)

        Try
            GraphicsBuilder.PageSelectObject(m_AnimationNumber)
            m_FILL_ON_ColorExpression = piValue

            GraphicsBuilder.PropertiesFillColourPut(m_FillColorType, m_FILL_ON_ColorExpression, "", "", "", "", False, 0, 0)
            'MessageBox.Show(GraphicsBuilder.PropertyVisibility.ToString())
            'GraphicsBuilder.PropertyVisibility.Remove '?????????????????????????????????????
            Exit Function
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try

    End Function

    Public Function SetLineColor(piValue As Integer)
        Try
            m_FILL_ON_ColorExpression = piValue
            GraphicsBuilder.PageSelectObject(m_AnimationNumber)

            GraphicsBuilder.AttributeLineColour = piValue
            Exit Function
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Function

    Public Function SetPropertyVisibility(text As String)
        Try
            GraphicsBuilder.PageSelectObject(m_AnimationNumber)
            GraphicsBuilder.PropertyVisibility = text
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Function


    Public Sub SetPipeColor(piLoLightColor As Integer, piHiLightColor As Integer)

        GraphicsBuilder.PageSelectObject(m_AnimationNumber)

        If LCase(m_CitectObjectType) = "pipe" Then
            GraphicsBuilder.AttributeHiLightColour = piLoLightColor
            GraphicsBuilder.AttributeLoLightColour = piHiLightColor
        End If

        Exit Sub
ERR_HANDLER:
        MsgBox(Err.Description)
    End Sub



    Private Sub Class_Initialize()
        Clear()
    End Sub

    Public Sub SelectObj()
        GraphicsBuilder.PageSelectObject(m_AnimationNumber)
    End Sub

    Public ReadOnly Property LibraryName() As String
        Get
            LibraryName = m_LibraryName
        End Get
    End Property

    Public ReadOnly Property SymbolName() As String
        Get
            SymbolName = m_SymbolName
        End Get
    End Property

    Private Sub Class_Terminate()
        MsgBox("wyf!")
    End Sub

End Class
