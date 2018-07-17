Imports Microsoft.Office.Interop.Excel


Public Class frmMacroParameters


    Private m_ExcelSheet As Worksheet
    ' Create new Application
    Dim Excel As Application = New Application()
    ' Open Excel spreadsheet
    Dim xlWorkbook As Workbook = Excel.Workbooks.Open("")
    ' Loop over all sheets
    'For i As Integer = 0 To w.Sheets.Count


    Public IOList As String
    Public Description As String
    Public MODBUS_ADDRESS As String
    Public SFI_NUMBER As String
    Public NODE As String
    Public AlarmNo As String
    Public Unit As String
    Public SheetNo As String
    Public System As String
    Public Alarm As String

    Public IASTagname As String
    Public ToolTip As String
    Public Label As String


    Public Cancelled As Boolean


    Private Sub cmdBrowseIOList_Click()
        Dim LFName As String
        LFName = m_ExcelSheet.ToString
        If Not LFName = "" And Not LFName = "False" Then txtIOLIst.Text = LFName
    End Sub



    Private Sub cmdCancel_Click()
        Cancelled = True
        Me.Hide()
    End Sub



    Private Sub cmdOK_Click()
        Dim lSheet As Worksheet
        lSheet = Excel.ActiveWorkbook.Sheets(1)


        IOList = txtIOLIst.Text
        lSheet.Cells(3, 2) = IOList

        Description = txtDescription.Text
        lSheet.Cells(4, 2) = Description

        MODBUS_ADDRESS = txtModbus.Text
        lSheet.Cells(5, 2) = MODBUS_ADDRESS

        SFI_NUMBER = txtSFI.Text
        lSheet.Cells(6, 2) = SFI_NUMBER

        AlarmNo = txtAlarmNo.Text
        lSheet.Cells(7, 2) = AlarmNo

        Unit = txtUnit.Text
        lSheet.Cells(8, 2) = Unit

        lSheet.Cells(9, 2) = txtSheet.Text
        SheetNo = txtSheet.Text

        lSheet.Cells(10, 2) = txtSystem.Text
        System = txtSystem.Text

        lSheet.Cells(11, 2) = txtIASTagname.Text
        IASTagname = txtIASTagname.Text

        lSheet.Cells(12, 2) = txtTooltip.Text
        ToolTip = txtTooltip.Text

        lSheet.Cells(13, 2) = txtLabel.Text
        Label = txtLabel.Text

        lSheet.Cells(14, 2) = txtIOLIst.Text
        Alarm = txtIOLIst.Text

        lSheet.Cells(15, 2) = txtNode.Text
        NODE = txtNode.Text

        Me.Hide()

    End Sub







    Private Sub UserForm_Activate()
        Cancelled = False
        Dim lSheet As Worksheet
        lSheet = Excel.ActiveWorkbook.Sheets("Configuration")



        txtIOLIst = lSheet.Cells(3, 2)
        IOList = txtIOLIst.Text

        txtDescription = lSheet.Cells(4, 2)
        Description = txtDescription.Text

        txtModbus = lSheet.Cells(5, 2)
        MODBUS_ADDRESS = txtModbus.Text

        txtSFI = lSheet.Cells(6, 2)
        SFI_NUMBER = txtSFI.Text

        txtAlarmNo = lSheet.Cells(7, 2)
        AlarmNo = txtAlarmNo.Text

        txtUnit = lSheet.Cells(8, 2)
        Unit = txtUnit.Text

        txtSheet = lSheet.Cells(9, 2)
        SheetNo = txtSheet.Text

        txtSystem = lSheet.Cells(10, 2)
        System = txtSystem.Text

        txtIASTagname = lSheet.Cells(11, 2)
        IASTagname = txtIASTagname.Text

        txtTooltip = lSheet.Cells(12, 2)
        ToolTip = txtTooltip.Text

        txtLabel = lSheet.Cells(13, 2)
        Label = txtLabel.Text

        txtIOLIst = lSheet.Cells(14, 2)
        Alarm = txtIOLIst.Text

        txtNode = lSheet.Cells(15, 2)
        NODE = txtNode.Text

    End Sub

    Private Sub InitializeComponent()
        Me.txtIOLIst = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.txtModbus = New System.Windows.Forms.TextBox()
        Me.txtSFI = New System.Windows.Forms.TextBox()
        Me.txtAlarmNo = New System.Windows.Forms.TextBox()
        Me.txtUnit = New System.Windows.Forms.TextBox()
        Me.txtSheet = New System.Windows.Forms.TextBox()
        Me.txtSystem = New System.Windows.Forms.TextBox()
        Me.txtIASTagname = New System.Windows.Forms.TextBox()
        Me.txtTooltip = New System.Windows.Forms.TextBox()
        Me.txtLabel = New System.Windows.Forms.TextBox()
        Me.txtNode = New System.Windows.Forms.TextBox()
        Me.cmdBrowseIOList = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtIOLIst
        '
        Me.txtIOLIst.Location = New System.Drawing.Point(83, 35)
        Me.txtIOLIst.Multiline = True
        Me.txtIOLIst.Name = "txtIOLIst"
        Me.txtIOLIst.Size = New System.Drawing.Size(317, 20)
        Me.txtIOLIst.TabIndex = 0
        Me.txtIOLIst.Text = "txtIOLIst"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(222, 114)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(89, 20)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Text = "txtAlarm"
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(222, 140)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(89, 20)
        Me.txtDescription.TabIndex = 0
        Me.txtDescription.Text = "txtDescription"
        '
        'txtModbus
        '
        Me.txtModbus.Location = New System.Drawing.Point(222, 166)
        Me.txtModbus.Name = "txtModbus"
        Me.txtModbus.Size = New System.Drawing.Size(89, 20)
        Me.txtModbus.TabIndex = 0
        Me.txtModbus.Text = "txtModbus"
        '
        'txtSFI
        '
        Me.txtSFI.Location = New System.Drawing.Point(222, 192)
        Me.txtSFI.Name = "txtSFI"
        Me.txtSFI.Size = New System.Drawing.Size(89, 20)
        Me.txtSFI.TabIndex = 0
        Me.txtSFI.Text = "txtSFI"
        '
        'txtAlarmNo
        '
        Me.txtAlarmNo.Location = New System.Drawing.Point(222, 218)
        Me.txtAlarmNo.Name = "txtAlarmNo"
        Me.txtAlarmNo.Size = New System.Drawing.Size(89, 20)
        Me.txtAlarmNo.TabIndex = 0
        Me.txtAlarmNo.Text = "txtAlarmNo"
        '
        'txtUnit
        '
        Me.txtUnit.Location = New System.Drawing.Point(222, 244)
        Me.txtUnit.Name = "txtUnit"
        Me.txtUnit.Size = New System.Drawing.Size(89, 20)
        Me.txtUnit.TabIndex = 0
        Me.txtUnit.Text = "txtUnit"
        '
        'txtSheet
        '
        Me.txtSheet.Location = New System.Drawing.Point(222, 270)
        Me.txtSheet.Multiline = True
        Me.txtSheet.Name = "txtSheet"
        Me.txtSheet.Size = New System.Drawing.Size(89, 20)
        Me.txtSheet.TabIndex = 0
        Me.txtSheet.Text = "txtSheet"
        '
        'txtSystem
        '
        Me.txtSystem.Location = New System.Drawing.Point(222, 296)
        Me.txtSystem.Name = "txtSystem"
        Me.txtSystem.Size = New System.Drawing.Size(89, 20)
        Me.txtSystem.TabIndex = 0
        Me.txtSystem.Text = "txtSystem"
        '
        'txtIASTagname
        '
        Me.txtIASTagname.Location = New System.Drawing.Point(222, 322)
        Me.txtIASTagname.Name = "txtIASTagname"
        Me.txtIASTagname.Size = New System.Drawing.Size(89, 20)
        Me.txtIASTagname.TabIndex = 0
        Me.txtIASTagname.Text = "txtIASTagname"
        '
        'txtTooltip
        '
        Me.txtTooltip.Location = New System.Drawing.Point(222, 348)
        Me.txtTooltip.Name = "txtTooltip"
        Me.txtTooltip.Size = New System.Drawing.Size(89, 20)
        Me.txtTooltip.TabIndex = 0
        Me.txtTooltip.Text = "txtTooltip"
        '
        'txtLabel
        '
        Me.txtLabel.Location = New System.Drawing.Point(222, 374)
        Me.txtLabel.Name = "txtLabel"
        Me.txtLabel.Size = New System.Drawing.Size(89, 20)
        Me.txtLabel.TabIndex = 0
        Me.txtLabel.Text = "txtLabel"
        '
        'txtNode
        '
        Me.txtNode.Location = New System.Drawing.Point(222, 400)
        Me.txtNode.Name = "txtNode"
        Me.txtNode.Size = New System.Drawing.Size(89, 20)
        Me.txtNode.TabIndex = 0
        Me.txtNode.Text = "txtNode"
        '
        'cmdBrowseIOList
        '
        Me.cmdBrowseIOList.Location = New System.Drawing.Point(406, 35)
        Me.cmdBrowseIOList.Name = "cmdBrowseIOList"
        Me.cmdBrowseIOList.Size = New System.Drawing.Size(24, 20)
        Me.cmdBrowseIOList.TabIndex = 3
        Me.cmdBrowseIOList.Text = "..."
        Me.cmdBrowseIOList.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(179, 466)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(82, 39)
        Me.cmdOK.TabIndex = 4
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(267, 466)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(82, 39)
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'frmMacroParameters
        '
        Me.ClientSize = New System.Drawing.Size(552, 532)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdBrowseIOList)
        Me.Controls.Add(Me.txtNode)
        Me.Controls.Add(Me.txtLabel)
        Me.Controls.Add(Me.txtTooltip)
        Me.Controls.Add(Me.txtIASTagname)
        Me.Controls.Add(Me.txtSystem)
        Me.Controls.Add(Me.txtSheet)
        Me.Controls.Add(Me.txtUnit)
        Me.Controls.Add(Me.txtAlarmNo)
        Me.Controls.Add(Me.txtSFI)
        Me.Controls.Add(Me.txtModbus)
        Me.Controls.Add(Me.txtDescription)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.txtIOLIst)
        Me.Name = "frmMacroParameters"
        Me.Text = "frmMacroParameters"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private Sub frmMacroParameters_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub txtAlarm_TextChanged(sender As Object, e As EventArgs) Handles txtIOLIst.TextChanged, txtIOLIst.TextChanged

    End Sub

    Private Sub txtAlarmNo_TextChanged(sender As Object, e As EventArgs) Handles txtAlarmNo.TextChanged

    End Sub
End Class