Imports Microsoft.Office.Interop.Excel

Module modGlobal

    Public GetMacroWorkBook As Workbook

    Public Function SetWorkBook(piWorkBook As Workbook)
        GetMacroWorkBook = piWorkBook
    End Function

    Public Function Configure()
        modConfiguration.showConfigDialog()
    End Function

End Module
