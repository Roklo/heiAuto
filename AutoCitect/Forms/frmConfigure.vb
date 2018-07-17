Imports Microsoft.Office.Interop.Excel



Public Class frmConfigure

    Private m_ExcelSheet As Worksheet
    ' Create new Application
    Dim Excel As Application = New Application()
    ' Open Excel spreadsheet
    Dim xlWorkbook As Workbook = Excel.Workbooks.Open("")
    ' Loop over all sheets
    'For i As Integer = 0 To w.Sheets.Count

    Private Sub frmConfigure_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public Overloads Sub ShowDialog()

        ColumnAlocation()
        Me.Show()
        m_ExcelSheet = Nothing
    End Sub

    Function ColumnAlocation()

        m_ExcelSheet = Excel.Sheets("Configuration")

        Me.txtTagName.Text = CStr(m_ExcelSheet.Cells(5, "C")) 'Tagname
        Me.txtDescription.Text = CStr(m_ExcelSheet.Cells(7, "C")) 'Signal Description

        Me.txtAlarmGroup.Text = CStr(m_ExcelSheet.Cells(9, "C")) 'Alarm      Group

        Me.txtBlock.Text = CStr(m_ExcelSheet.Cells(10, "C")) 'Block

        Me.txtAlarmDelay.Text = CStr(m_ExcelSheet.Cells(13, "C")) 'Alarm     Delay
        Me.txtLLAlarmLimit.Text = CStr(m_ExcelSheet.Cells(14, "C")) 'Allim         LL
        Me.txtLAlarmLimit.Text = CStr(m_ExcelSheet.Cells(15, "C")) 'Allim       L
        Me.txtHAlarmLimit.Text = CStr(m_ExcelSheet.Cells(16, "C")) 'Allim         H
        Me.txtHHAlarmLimit.Text = CStr(m_ExcelSheet.Cells(17, "C")) 'Allim         HH
        Me.txtNode.Text = CStr(m_ExcelSheet.Cells(18, "C")) 'Node
        Me.txtStationNo.Text = CStr(m_ExcelSheet.Cells(19, "C")) 'Station No.
        Me.txtIOCardNo.Text = CStr(m_ExcelSheet.Cells(20, "C")) 'IO-Card No.
        Me.txtChannel.Text = CStr(m_ExcelSheet.Cells(21, "C")) 'Channel
        Me.txtRawMin.Text = CStr(m_ExcelSheet.Cells(22, "C")) 'RAW Min
        Me.txtRawMax.Text = CStr(m_ExcelSheet.Cells(23, "C")) 'RAW Max
        Me.txtEngMin.Text = CStr(m_ExcelSheet.Cells(24, "C")) 'IAS-RangeMin
        Me.txtEngMax.Text = CStr(m_ExcelSheet.Cells(25, "C")) 'IAS-RangeMax
        Me.txtAlarmNo.Text = CStr(m_ExcelSheet.Cells(26, "C")) 'IAS-Tagname
        Me.txtAddress.Text = CStr(m_ExcelSheet.Cells(27, "C")) 'Adress
        Me.txtAddress2.Text = CStr(m_ExcelSheet.Cells(29, "C")) 'Address2
        Me.txtModbusAddress.Text = CStr(m_ExcelSheet.Cells(28, "C")) 'Modbus Address
        Me.txtAlarmYesNo.Text = CStr(m_ExcelSheet.Cells(8, "C")) 'Alarm Y/N
        Me.txtUnit.Text = CStr(m_ExcelSheet.Cells(12, "C")) 'Unit
        Me.txtSignalType.Text = CStr(m_ExcelSheet.Cells(11, "C")) 'Signal Type
        Me.txtISA.Text = CStr(m_ExcelSheet.Cells(4, "C")) 'ISA
        Me.txtTagNo.Text = CStr(m_ExcelSheet.Cells(3, "C")) 'Tag No.
        Me.txtSystemName.Text = CStr(m_ExcelSheet.Cells(6, "C")) 'System Name
        Me.txtNormalValue.Text = CStr(m_ExcelSheet.Cells(30, "C")) 'Normal Value
        Me.txtStartAlarmPS.Text = CStr(m_ExcelSheet.Cells(31, "C")) 'Start Alarm No PS
        Me.txtStartAlarmSB.Text = CStr(m_ExcelSheet.Cells(32, "C")) 'Start Alarm No SB
        Me.txtAkerDBF.Text = CStr(m_ExcelSheet.Cells(33, "C")) 'Path To AKER.DBF


        'New 09 Feb 2010
        Me.txtLabel.Text = CStr(m_ExcelSheet.Cells(34, "C"))
        Me.txtToolTip.Text = CStr(m_ExcelSheet.Cells(35, "C"))
        Me.txtPathIO_List.Text = CStr(m_ExcelSheet.Cells(36, "C"))
        Me.txtIOListSheet.Text = CStr(m_ExcelSheet.Cells(37, "C"))


        modConfiguration.TNO = CStr(m_ExcelSheet.Cells(3, "C")) 'Tag No.
        modConfiguration.ISA = CStr(m_ExcelSheet.Cells(4, "C")) 'ISA
        modConfiguration.TN = CStr(m_ExcelSheet.Cells(5, "C")) 'Tag Name
        modConfiguration.SYSN = CStr(m_ExcelSheet.Cells(6, "C")) 'System Name
        modConfiguration.DES = CStr(m_ExcelSheet.Cells(7, "C")) 'Signal Description
        modConfiguration.AYN = CStr(m_ExcelSheet.Cells(8, "C")) 'Alarm
        modConfiguration.AG = CStr(m_ExcelSheet.Cells(9, "C")) 'Alarm      Group
        modConfiguration.BG = CStr(m_ExcelSheet.Cells(10, "C")) 'Block
        modConfiguration.ST = CStr(m_ExcelSheet.Cells(11, "C")) 'Signal Type
        modConfiguration.UN = CStr(m_ExcelSheet.Cells(12, "C")) 'Unit
        modConfiguration.ALD = CStr(m_ExcelSheet.Cells(13, "C")) 'Alarm     Delay
        modConfiguration.ALL = CStr(m_ExcelSheet.Cells(14, "C")) 'Allim         LL
        modConfiguration.AL = CStr(m_ExcelSheet.Cells(15, "C")) 'Allim       L
        modConfiguration.AH = CStr(m_ExcelSheet.Cells(16, "C")) 'Allim         H
        modConfiguration.AHH = CStr(m_ExcelSheet.Cells(17, "C")) 'Allim         HH
        modConfiguration.NODE = CStr(m_ExcelSheet.Cells(18, "C")) 'Node
        modConfiguration.SN = CStr(m_ExcelSheet.Cells(19, "C")) 'Station No.
        modConfiguration.IOCNO = CStr(m_ExcelSheet.Cells(20, "C")) 'IO-Card No.
        modConfiguration.CH = CStr(m_ExcelSheet.Cells(21, "C")) 'Channel
        modConfiguration.RMin = CStr(m_ExcelSheet.Cells(22, "C")) 'RAW Min
        modConfiguration.RMax = CStr(m_ExcelSheet.Cells(23, "C")) 'RAW Max
        modConfiguration.Emin = CStr(m_ExcelSheet.Cells(24, "C")) 'IAS-RangeMin
        modConfiguration.EMax = CStr(m_ExcelSheet.Cells(25, "C")) 'IAS-RangeMax
        modConfiguration.IasTN = CStr(m_ExcelSheet.Cells(26, "C")) 'IAS-Tagname
        modConfiguration.Adr = CStr(m_ExcelSheet.Cells(27, "C")) 'Adress
        modConfiguration.MbAdd = CStr(m_ExcelSheet.Cells(28, "C")) 'Modbus Address
        modConfiguration.Adr2 = CStr(m_ExcelSheet.Cells(29, "C")) 'Address2
        modConfiguration.NVAL = CStr(m_ExcelSheet.Cells(30, "C")) 'Normal Value
        modConfiguration.SANP = CStr(m_ExcelSheet.Cells(31, "C")) 'Start Alarm No PS
        modConfiguration.SANS = CStr(m_ExcelSheet.Cells(32, "C")) 'Start Alarm No SB
        modConfiguration.Aker_PATH = CStr(m_ExcelSheet.Cells(33, "C")) 'Path To AKER.DBF

        modConfiguration.ColumnAlloc_Label = CStr(m_ExcelSheet.Cells(34, "C")) '
        modConfiguration.ColumnAlloc_ToolTip = CStr(m_ExcelSheet.Cells(35, "C")) '
        modConfiguration.IOList_Path = CStr(m_ExcelSheet.Cells(36, "C")) '
        modConfiguration.IOLIST_Sheet = CStr(m_ExcelSheet.Cells(37, "C")) '
        modConfiguration.MaxAlm = CStr(m_ExcelSheet.Cells(38, "C")) '

    End Function
    Function SaveColumnAlocation()


        m_ExcelSheet.Cells(3, "C") = txtTagNo
        m_ExcelSheet.Cells(4, "C") = txtISA
        m_ExcelSheet.Cells(5, "C") = txtTagName
        m_ExcelSheet.Cells(6, "C") = txtSystemName
        m_ExcelSheet.Cells(7, "C") = txtDescription
        m_ExcelSheet.Cells(8, "C") = txtAlarmYesNo
        m_ExcelSheet.Cells(9, "C") = txtAlarmGroup
        m_ExcelSheet.Cells(10, "C") = txtBlock
        m_ExcelSheet.Cells(11, "C") = txtSignalType
        m_ExcelSheet.Cells(12, "C") = txtUnit
        m_ExcelSheet.Cells(13, "C") = txtAlarmDelay
        m_ExcelSheet.Cells(14, "C") = txtLLAlarmLimit
        m_ExcelSheet.Cells(15, "C") = txtLAlarmLimit
        m_ExcelSheet.Cells(16, "C") = txtHAlarmLimit
        m_ExcelSheet.Cells(17, "C") = txtHHAlarmLimit
        m_ExcelSheet.Cells(18, "C") = txtNode
        m_ExcelSheet.Cells(19, "C") = txtStationNo
        m_ExcelSheet.Cells(20, "C") = txtIOCardNo
        m_ExcelSheet.Cells(21, "C") = txtChannel
        m_ExcelSheet.Cells(22, "C") = txtRawMin
        m_ExcelSheet.Cells(23, "C") = txtRawMax
        m_ExcelSheet.Cells(24, "C") = txtEngMin
        m_ExcelSheet.Cells(25, "C") = txtEngMax
        m_ExcelSheet.Cells(26, "C") = txtAlarmNo
        m_ExcelSheet.Cells(27, "C") = txtAddress
        m_ExcelSheet.Cells(28, "C") = txtModbusAddress
        m_ExcelSheet.Cells(29, "C") = txtAddress2
        m_ExcelSheet.Cells(30, "C") = txtNormalValue
        m_ExcelSheet.Cells(31, "C") = txtStartAlarmPS
        m_ExcelSheet.Cells(32, "C") = txtStartAlarmSB
        m_ExcelSheet.Cells(33, "C") = Me.txtAkerDBF


        'New 09 Feb 2010
        m_ExcelSheet.Cells(34, "C") = Me.txtLabel
        m_ExcelSheet.Cells(35, "C") = Me.txtToolTip
        m_ExcelSheet.Cells(36, "C") = Me.txtPathIO_List
        m_ExcelSheet.Cells(37, "C") = Me.txtIOListSheet

        ColumnAlocation()
    End Function



    Private Sub cmdBrowseForIOList_Click()
        Dim LFName As String
        LFName = Excel.GetOpenFilename
        If Not LFName = "" And Not LFName = "False" Then txtPathIO_List.Text = LFName
    End Sub

    Public Sub LoadConfiguration()
        ColumnAlocation()
    End Sub

    Private Sub cmdExit_Click()
        Me.Hide()
    End Sub

    Private Sub cmdSave_Click()
        SaveColumnAlocation()
    End Sub

    Private Sub Label3_Click()

    End Sub

    Private Sub UserForm_Click()

    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles txtHHAlarmLimit.TextChanged

    End Sub

    Private Sub txtBlock_TextChanged(sender As Object, e As EventArgs) Handles txtBlock.TextChanged

    End Sub

    Private Sub textbox12_TextChanged(sender As Object, e As EventArgs)

    End Sub
End Class