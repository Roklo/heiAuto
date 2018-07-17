Option Explicit On
Imports GraphicsBuilder
Imports Microsoft.VisualBasic



Public Class clsCITECTData

    Dim GraphicsBuilder As GraphicsBuilderClass = New GraphicsBuilderClass()
    Dim counter As Integer = 0

    Private m_CitectObjects As Collection
    Private m_Genies As Collection
    Private m_ProjectName As String


    Private Sub Clear()
        m_CitectObjects = New Collection
        m_Genies = New Collection
        m_ProjectName = ""
    End Sub

    Public ReadOnly Property ObjectCount() As Long
        Get
            Try
                ObjectCount = m_CitectObjects.Count
            Catch ex As Exception
                ObjectCount = 0
                Dim testmsg As Integer
                testmsg = MsgBox("Error: Was not able to get ObjectCount. Make sure Citect Scada Graphics builder is running and the mimic you want to edit is open",
                                 vbOKOnly + vbCritical, "Error")
                Exit Property
            End Try
        End Get

    End Property

    Public Function GetObject(piIndex As Long) As clsCITECTObject
        On Error Resume Next
        GetObject = m_CitectObjects(piIndex)
    End Function

    Public ReadOnly Property GenieCount() As Long
        Get
            On Error Resume Next
            GenieCount = m_Genies.Count
        End Get


    End Property

    Public Function GetGenie(piIndex As Long) As clsCITECTGenie
        On Error Resume Next
        GetGenie = m_Genies(piIndex)
    End Function


    Public ReadOnly Property GetCitectObjectType() As String
        Get

            Try
                GetCitectObjectType = GraphicsBuilder.AttributeClass()
            Catch ex As Exception
                GetCitectObjectType = ""
                Dim testmsg As Integer
                testmsg = MsgBox("Error: Was not able to get CitectObjectType. Make sure Citect Scada Graphics builder is running and the mimic you want to edit is open",
                                 vbOKOnly + vbCritical, "Error")
                Exit Property
            End Try

        End Get

    End Property

    Public Property GetIASGenieType() As String
        Get
            'Do nothing
        End Get
        Set(value As String)
            'Do nothing
        End Set
    End Property

    Public ReadOnly Property GetAnimationNumber() As Short
        Get
            On Error GoTo ERR_HANDLER
            GetAnimationNumber = GraphicsBuilder.AttributeAN
            counter = counter + 1


            'MessageBox.Show("Fant:" + GetAnimationNumber.ToString)


            'If GetAnimationNumber = Nothing Then
            ' MessageBox.Show("Fant ikke noe." + GetAnimationNumber.ToString)
            'Else
            'MessageBox.Show("Fant: " + GetAnimationNumber.ToString)
            'End If

ERR_HANDLER:
            GetAnimationNumber = GetAnimationNumber
            'Dim testmsg As Integer
            'testmsg = MsgBox("Error: Was not able to get AnimationNumber. Make sure Citect Scada Graphics builder is running and the mimic you want to edit is open",
            'vbOKOnly + vbCritical, "Error")
            Exit Property
        End Get
    End Property




    Private Sub Class_Initialize()
        Clear()
    End Sub

    Public Function PageSelectFirstObject() As Boolean
        Try

            Clear()
            GraphicsBuilder.PageSelectFirstObject()
            PageSelectFirstObject = True

        Catch ex As Exception
            PageSelectFirstObject = False
            Exit Function
        End Try


    End Function

    Public Function PageSelectNextObject() As Boolean
        Try

            GraphicsBuilder.PageSelectNextObject()
            PageSelectNextObject = True
        Catch ex As Exception
            PageSelectNextObject = False
            Exit Function
        End Try

    End Function

    Private Function ProcessGenie()
        Dim lGenie As clsCITECTGenie
        lGenie = New clsCITECTGenie
        m_Genies.Add(lGenie)
        lGenie.ParseGenie()

    End Function



    Public Function ParsePage()
        Dim lAnim As Short
        Dim lStat As Boolean
        Dim lGenie As clsCITECTGenie
        Dim lObj As clsCITECTObject

        m_ProjectName = GraphicsBuilder.ProjectSelected
        lStat = PageSelectFirstObject()

        While lStat

            lAnim = GetAnimationNumber
            'MessageBox.Show("lAnim: " + lAnim.ToString)
            If lAnim > 0 Then

                If GetCitectObjectType = "Genie" Then

                    lGenie = New clsCITECTGenie
                    m_Genies.Add(lGenie)
                    'Debug.Print("gENIE cOUNT = " + CStr(m_Genies.Count()))
                    lGenie.ParseGenie()
                    lAnim = lAnim
                    'MessageBox.Show("Added Genie object.")
                ElseIf GetCitectObjectType = "Line" Then
                    lObj = New clsCITECTObject
                    m_CitectObjects.Add(lObj)
                    lObj.ParseObject()
                    'MessageBox.Show("Added Line object.")
                ElseIf GetCitectObjectType = "Pipe" Then
                    lObj = New clsCITECTObject
                    m_CitectObjects.Add(lObj)
                    lObj.ParseObject()
                    'MessageBox.Show("Added Pipe object.")
                ElseIf GetCitectObjectType = "CircleV2" Then '??ELLIPSE**
                    lObj = New clsCITECTObject
                    m_CitectObjects.Add(lObj)
                    lObj.ParseObject()
                    'MessageBox.Show("Added CircleV2 object.")
                Else
                    lObj = New clsCITECTObject
                    m_CitectObjects.Add(lObj)
                    lObj.ParseObject()
                    'MessageBox.Show("Added some object.")
                End If

                'MessageBox.Show(GetAnimationNumber.ToString + " " + lAnim.ToString)
                'MessageBox.Show("Added " + lObj.ToString + "." + "Total objects: " + m_CitectObjects.Count.ToString)

            End If
            lStat = PageSelectNextObject()
        End While

        'MessageBox.Show("hei " + counter.ToString + " " + m_CitectObjects.Count.ToString)

    End Function

    'Searches for genies by name
    Public Function FindGeniesBy(PiPropertyName As String, piPropertyValue As String) As Collection
        Dim i As Long

        Dim lRetval As Collection
        lRetval = New Collection


        For i = 1 To GenieCount
            Dim lGenie As clsCITECTGenie
            lGenie = GetGenie(i)
            If lGenie.GetPropertyValueByName(PiPropertyName) = piPropertyValue Then
                lRetval.Add(lGenie)
            End If
        Next

        FindGeniesBy = lRetval
    End Function

    Public Function GetGeniePropertyNames() As clsCitectPropertyNames
        Dim i As Long
        Dim lGen As clsCITECTGenie
        Dim lNames As clsCitectPropertyNames
        lNames = New clsCitectPropertyNames

        For i = 1 To GenieCount
            lGen = Me.GetGenie(i)
            Dim Lprop As clsCITECTProperty

            Dim j As Long
            For j = 1 To lGen.PropertyCount
                Lprop = lGen.GetProperty(j)
                lNames.Add(Lprop)
            Next
        Next

        GetGeniePropertyNames = lNames
    End Function

End Class

