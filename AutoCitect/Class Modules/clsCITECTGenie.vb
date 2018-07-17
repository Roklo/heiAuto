Option Explicit On
Imports GraphicsBuilder


Public Class clsCITECTGenie
    Dim GraphicsBuilder As GraphicsBuilderClass = New GraphicsBuilderClass()
    Private m_AnimationNumber As Integer
    Private m_IsTagged As Boolean
    Private m_LibraryName As String
    Private m_ObjectName As String
    Private m_ObjectType As Integer
    Private m_ToolTip As String


    Private m_IASGenieType As String
    Private m_SFI_NUMBER As String

    'Private m_CitectObjects As Collection
    Private m_Properties As Collection

    Public ReadOnly Property LibraryName() As String
        Get
            LibraryName = m_LibraryName
        End Get
    End Property

    Public ReadOnly Property Tooltip() As String
        Get
            Tooltip = m_ToolTip
        End Get
    End Property

    Public ReadOnly Property ObjectName() As String
        Get
            ObjectName = m_ObjectName
        End Get
    End Property


    Private Sub Clear()
        ' m_AnimationNumber = -1

        '  m_IASGenieType = ""
        ' m_SFI_NUMBER = ""

        'Set m_CitectObjects = New Collection
        m_Properties = New Collection
    End Sub

    'Public Property Get ObjectCount() As Long
    '    On Error Resume Next
    '    ObjectCount = m_CitectObjects.Count
    'End Property

    'Public Function GetObject(piIndex As Long) As clsCITECTObject
    '    On Error Resume Next
    '    Set GetObject = m_CitectObjects(piIndex)
    'End Function

    Public ReadOnly Property PropertyCount() As Long
        Get
            On Error Resume Next
            PropertyCount = m_Properties.Count
        End Get
    End Property

    Public Function GetProperty(piIndex As Long) As clsCITECTProperty
        On Error Resume Next
        GetProperty = m_Properties(piIndex)
    End Function


    'case insensitive lookup function
    Public Function GetPropertyByName(piName As String) As clsCITECTProperty
        On Error GoTo ERROR_HANDLER

        Dim i As Long
        Dim lRetval As clsCITECTProperty

        piName = Trim(piName) 'remove whitespace


        For i = 1 To PropertyCount()
            lRetval = m_Properties(i)
            If Not lRetval Is Nothing Then
                If LCase(lRetval.Name) = LCase(piName) Then
                    GetPropertyByName = m_Properties(i)
                    Exit For
                End If
            End If
        Next

        '*****HACK Due to misc variations in name****
        If piName = "AlarmNo" And GetPropertyByName Is Nothing Then
            piName = "Alarm No"
            GetPropertyByName = GetPropertyByName(piName)
        End If
        '---------------------------------------------


        Exit Function
ERROR_HANDLER:
        MsgBox(Err.Description)
    End Function

    'case insensitive(on piPropertyName) update function
    Public Function UpdateProperty(PiPropertyName As String, piValue As String)


        Try
            Dim Lprop As clsCITECTProperty
            Lprop = GetPropertyByName(PiPropertyName)

            If LCase(PiPropertyName) = "alarmno" Then

                m_AnimationNumber = m_AnimationNumber

            End If

            If (LCase(PiPropertyName) = "alarmno" Or LCase(PiPropertyName) = "alarm no" Or LCase(PiPropertyName) = "alarm no:") And Lprop Is Nothing Then
                PiPropertyName = "Alarm No"
                Lprop = GetPropertyByName("Alarm No")
            End If

            If (LCase(PiPropertyName) = "alarmno" Or LCase(PiPropertyName) = "alarm no" Or LCase(PiPropertyName) = "alarm no:") And Lprop Is Nothing Then
                PiPropertyName = "Alarm No:"
                Lprop = GetPropertyByName("Alarm No:")
            End If

            If (LCase(PiPropertyName) = "alarmno" Or LCase(PiPropertyName) = "alarm no" Or LCase(PiPropertyName) = "alarm no:") And Lprop Is Nothing Then
                PiPropertyName = "AlarmNo:"
                Lprop = GetPropertyByName("AlarmNo:")
            End If


            If Lprop Is Nothing Then
                Dim testmsg As Integer
                testmsg = MsgBox("Error: Did not find property", vbOKOnly + vbCritical, "Error")
                Exit Function
            End If

            Lprop.Value = piValue

            GraphicsBuilder.PageSelectObject(m_AnimationNumber)
            GraphicsBuilder.LibraryObjectPutProperty(PiPropertyName, piValue)

            Exit Function
        Catch ex As Exception
            'MsgBox Err.DESCRIPTION
        End Try
    End Function

    Public Sub DeleteFromCitect()

        On Error Resume Next
        GraphicsBuilder.PageSelectObject(m_AnimationNumber)
        GraphicsBuilder.PageDeleteObject()

    End Sub

    Public Sub SelectObject()

        On Error Resume Next
        GraphicsBuilder.PageSelectObject(m_AnimationNumber)
    End Sub



    Public Function GetPropertyValueByName(piName As String) As String
        On Error Resume Next
        Dim Lprop As clsCITECTProperty
        Lprop = GetPropertyByName(piName)
        GetPropertyValueByName = Lprop.Value
    End Function

    Public Sub ParseGenie()


        Dim lStat As Boolean
        Dim lDesc As String
        Dim lAnimationNumber As Integer
        Dim Lobject As clsCITECTObject
        Dim lProjectName As String
        Dim lTooltip As String

        lProjectName = GraphicsBuilder.ProjectSelected


        m_AnimationNumber = GetAnimationNumber
        GraphicsBuilder.LibraryObjectName(lProjectName, m_LibraryName, m_ObjectName, m_ObjectType)


        'Debug.Print m_ObjectName

        GetGeineProperties()



    End Sub



    Public ReadOnly Property Get_SFI_NUMBER() As String
        Get
            If m_SFI_NUMBER = "" Then
                Dim Lprop As clsCITECTProperty
                Lprop = GetPropertyByName("SFI_NUMBER")
                If Not Lprop Is Nothing Then m_SFI_NUMBER = Lprop.Value
            End If
            Get_SFI_NUMBER = m_SFI_NUMBER
        End Get


    End Property


    Public ReadOnly Property GetCitectObjectType() As String
        Get
            GetCitectObjectType = "Genie"
        End Get

    End Property



    Private ReadOnly Property GetAnimationNumber() As Integer
        Get
            Try
                GetAnimationNumber = GraphicsBuilder.AttributeAN

                Exit Property
            Catch ex As Exception
                Clear()
                GetAnimationNumber = 0
            End Try
        End Get
    End Property

    Public ReadOnly Property GetAnimationNumber2() As Integer
        Get
            GetAnimationNumber2 = m_AnimationNumber
        End Get



    End Property


    Private Sub Class_Initialize()
        Clear()
        m_AnimationNumber = GetAnimationNumber
    End Sub



    Private Function PageSelectFirstObjectInGenie() As Boolean
        On Error GoTo ERR_HANDLER

        GraphicsBuilder.PageSelectFirstObjectInGenie()
        PageSelectFirstObjectInGenie = True
        Exit Function
ERR_HANDLER:
        PageSelectFirstObjectInGenie = False
    End Function

    Private Function PageSelectNextObjectInGenie() As Boolean
        On Error GoTo ERR_HANDLER
        GraphicsBuilder.PageSelectNextObjectInGenie()
        PageSelectNextObjectInGenie = True
        Exit Function
ERR_HANDLER:
        PageSelectNextObjectInGenie = False
    End Function


    Private Sub Class_Terminate()
        'Not in use?
    End Sub

    'This method gets the first propperty of the selected object in the Graphics Builder
    Private Function LibraryObjectFirstProperty() As clsCITECTProperty
        Try
            Dim lName As String
            Dim lval As String


            GraphicsBuilder.LibraryObjectFirstProperty(lName, lval)

            Dim lRetval As clsCITECTProperty
            lRetval = New clsCITECTProperty
            lRetval.Name = lName
            lRetval.Value = lval
            If lval IsNot ("") Then
                m_IsTagged = True
            End If
            LibraryObjectFirstProperty = lRetval

            Exit Function
        Catch ex As Exception

        End Try

    End Function

    'This method gets the next property of an selected object in the Graphics Builder
    Private Function LibraryObjectNextProperty() As clsCITECTProperty
        Try
            Dim lName As String
            Dim lval As String

            GraphicsBuilder.LibraryObjectNextProperty(lName, lval)

            Dim lRetval As clsCITECTProperty
            lRetval = New clsCITECTProperty
            lRetval.Name = lName
            lRetval.Value = lval
            If lval IsNot ("") Then
                m_IsTagged = True
            End If
            LibraryObjectNextProperty = lRetval
            Exit Function
        Catch ex As Exception
            ' MessageBox.Show("Could not get next property or there ar no more properties")
            'Do nothing
        End Try
    End Function


    'This method gets all the properties of an selected object in the Graphics Builder
    Private Sub GetGeineProperties()

        If PropertyCount > 0 Then
            Dim testmsg As Integer
            testmsg = MsgBox("Error: GetGeineProperties, called more then once for same genie",
                                 vbOKOnly + vbCritical, "Error")
        End If

        Dim Lprop As clsCITECTProperty = New clsCITECTProperty()
        Try
            Clear()
            Lprop = LibraryObjectFirstProperty()

            m_Properties.Add(Lprop)


        Catch ex As Exception
            MessageBox.Show("There was no first object")
            m_IsTagged = False
        End Try

        'Checks if the object infact has a next property
        Dim hasNextProperty As Boolean = True
        Try
            Dim lName
            Dim Lval
            GraphicsBuilder.LibraryObjectNextProperty(lName, Lval)
            If (lName Or Lval) Is Nothing Then
                hasNextProperty = False
            End If
        Catch ex As Exception
            hasNextProperty = False
        End Try


        If hasNextProperty = True Then
                While Not Lprop Is Nothing
                    Clear()
                    Lprop = LibraryObjectNextProperty()
                    m_Properties.Add(Lprop)

                End While
            End If

            Try
            If Lprop.Name.ToString.Equals("TOOLTIP") Then
                m_ToolTip = Lprop.Value
            End If
        Catch ex As Exception
            Exit Try
        End Try

        'm_AnimationNumber
        'm_IsTagged
        'm_LibraryName
        'm_ObjectName
        'm_ObjectType




    End Sub

    Public Function MarkAsTagged(piTagged As Boolean)
        m_IsTagged = piTagged
    End Function

    Public ReadOnly Property IsTagged() As Boolean
        Get
            IsTagged = m_IsTagged
        End Get
    End Property

End Class
