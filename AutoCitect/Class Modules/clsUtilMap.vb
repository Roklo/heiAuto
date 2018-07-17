Option Explicit On

Public Class clsUtilMap
    Private m_Map As Collection
    Private m_itemByIndex As Collection


    Private Sub Class_Initialize()
        m_Map = New Collection
    End Sub


    Public Function Add(piKey As String, piVal As String) As clsUtilMapEntry
        On Error Resume Next

        m_itemByIndex = Nothing

        Dim Litem As clsUtilMapEntry

        Litem = GetItem(piKey)
        If Litem Is Nothing Then
            Litem = New clsUtilMapEntry
            Litem.Key = piKey
            Litem.Value = piVal
            m_Map.Add(Litem, piKey)
        End If


        Litem.ItemCount = Litem.ItemCount + 1

    End Function

    Public Function ValExists(piKey As String) As Boolean
        On Error GoTo ERR_EXIT

        ValExists = True

        Dim lval As clsUtilMapEntry
        lval = m_Map(piKey)

        Exit Function
ERR_EXIT:
        ValExists = False
    End Function

    Public Function GetItem(piKey As String) As clsUtilMapEntry
        On Error Resume Next
        GetItem = m_Map(piKey)
    End Function

    Public Function GetKeyCount(piKey As String) As Long
        On Error Resume Next

        'GetKeyCount = m_Counter(piKey)
    End Function


    Public Function GetItemCount() As Long
        GetItemCount = m_Map.Count
    End Function

    Public Function GetItemByIndex(piIndex As Long) As clsUtilMapEntry
        If m_itemByIndex Is Nothing Then
            m_itemByIndex = New Collection
            Dim Litem As clsUtilMapEntry
            For Each Litem In m_Map
                m_itemByIndex.Add(Litem)
            Next
        End If


        GetItemByIndex = m_itemByIndex(piIndex)

    End Function


End Class
