Option Explicit On
Imports GraphicsBuilder
Imports Microsoft.Office.Interop.Excel




Public Class clsCitectLibrary
    Dim ExcelSheet As Worksheet
    Dim xlWorkbook = Nothing
    Private Excel As Application = New Application()

    Dim GraphicsBuilder As GraphicsBuilderClass = New GraphicsBuilderClass()
    Private m_LibraryNames As Collection


    Private Sub Class_Initialize()
        m_LibraryNames = New Collection
        Dim i As Long
        Dim lConfigSheet As Worksheet
        'Dim xxlWorkbook As Excel.Workbook
        'ThisWorkbook.Worksheets("Configuration")
        lConfigSheet = Excel.Sheets("Configuration")

        i = 3
        Dim lLibName As String
        lLibName = lConfigSheet.Cells(i, "F")
        While lLibName <> ""
            m_LibraryNames.Add(lLibName)
            i = i + 1
            lLibName = lConfigSheet.Cells(i, "F")
        End While
    End Sub


    Public Function PasteGenie(piProjectName As String, piGenieName As String, piXPos As Long, piYPos As Long, piSignal As IO_List_Signal, ByRef poXExtent As Long, poYExtent As Long) As Boolean
        PasteGenie = False

        poXExtent = 0
        poYExtent = 0


        Dim lGenie As clsCITECTGenie
        lGenie = tryPasteGenie_All_Libs(piProjectName, piGenieName, piXPos, piYPos)

        If lGenie Is Nothing Then Exit Function


        Dim i As Long
        Dim LPropVal As String
        Dim lPropName As String
        Dim Lprop As clsCITECTProperty

        'update all the properties of the genie
        For i = 1 To lGenie.PropertyCount
            Lprop = lGenie.GetProperty(i)
            LPropVal = Lprop.Value
            lPropName = Lprop.Name

            LPropVal = piSignal.GetPropertByName(lPropName)

            lGenie.UpdateProperty(lPropName, LPropVal)
        Next

        poXExtent = GraphicsBuilder.AttributeExtentX - piXPos
        poYExtent = GraphicsBuilder.AttributeExtentY - piYPos

        PasteGenie = True
    End Function


    Private Function tryPasteGenie_All_Libs(piProjectName As String, piGenieName As String, piXPos As Long, piYPos As Long) As clsCITECTGenie
        Dim lLibName As Object
        Dim lLibraryName As String
        Dim lGenie As clsCITECTGenie
        For Each lLibName In m_LibraryNames
            lLibraryName = lLibName
            lGenie = tryPasteGenie(piProjectName, lLibraryName, piGenieName, piXPos, piYPos)
            If Not lGenie Is Nothing Then
                Exit For
            End If
        Next
        tryPasteGenie_All_Libs = lGenie

    End Function

    Private Function tryPasteGenie(piProjectName As String, piLibraryName As String, piGenieName As String, piXPos As Long, piYPos As Long) As clsCITECTGenie
        Try
            GraphicsBuilder.LibraryObjectPlaceEx(piProjectName, piLibraryName, piGenieName, 1, True, piXPos, piYPos)

            Dim lRetval As clsCITECTGenie
            lRetval = New clsCITECTGenie
            lRetval.ParseGenie()

            tryPasteGenie = lRetval

            tryPasteGenie = lRetval

            Exit Function
        Catch ex As Exception
            tryPasteGenie = Nothing
        End Try
    End Function


End Class
