<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AutoCitect
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.cmdLoadFile = New System.Windows.Forms.Button()
        Me.cmdTestExcel = New System.Windows.Forms.Button()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.IOSetup = New System.Windows.Forms.Label()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.cmdFilePath = New System.Windows.Forms.TextBox()
        Me.lblFilePath = New System.Windows.Forms.Label()
        Me.lblTestPath = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.cmdBack = New System.Windows.Forms.Button()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdLoadFile
        '
        Me.cmdLoadFile.Location = New System.Drawing.Point(366, 51)
        Me.cmdLoadFile.Name = "cmdLoadFile"
        Me.cmdLoadFile.Size = New System.Drawing.Size(93, 23)
        Me.cmdLoadFile.TabIndex = 0
        Me.cmdLoadFile.Text = "Load Excel File"
        Me.cmdLoadFile.UseVisualStyleBackColor = True
        '
        'cmdTestExcel
        '
        Me.cmdTestExcel.Location = New System.Drawing.Point(365, 96)
        Me.cmdTestExcel.Name = "cmdTestExcel"
        Me.cmdTestExcel.Size = New System.Drawing.Size(93, 23)
        Me.cmdTestExcel.TabIndex = 0
        Me.cmdTestExcel.Text = "Test Excel"
        Me.cmdTestExcel.UseVisualStyleBackColor = True
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Location = New System.Drawing.Point(274, 164)
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(49, 20)
        Me.NumericUpDown1.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(227, 167)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Column:"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(169, 233)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(111, 42)
        Me.Button3.TabIndex = 0
        Me.Button3.Text = "Execute"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(104, 136)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(81, 17)
        Me.CheckBox1.TabIndex = 5
        Me.CheckBox1.Text = "CheckBox1"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(240, 136)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(81, 17)
        Me.CheckBox2.TabIndex = 5
        Me.CheckBox2.Text = "CheckBox1"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Location = New System.Drawing.Point(377, 136)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(81, 17)
        Me.CheckBox3.TabIndex = 5
        Me.CheckBox3.Text = "CheckBox1"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'IOSetup
        '
        Me.IOSetup.AutoSize = True
        Me.IOSetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IOSetup.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.IOSetup.Location = New System.Drawing.Point(232, 9)
        Me.IOSetup.Name = "IOSetup"
        Me.IOSetup.Size = New System.Drawing.Size(86, 20)
        Me.IOSetup.TabIndex = 6
        Me.IOSetup.Text = "I/O Setup"
        Me.IOSetup.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(207, 192)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(82, 23)
        Me.Button4.TabIndex = 7
        Me.Button4.Text = "Form_Load"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(292, 192)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(82, 23)
        Me.Button5.TabIndex = 7
        Me.Button5.Text = "PasteSymbol"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(377, 192)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(82, 23)
        Me.Button6.TabIndex = 7
        Me.Button6.Text = "PageSaved"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(286, 233)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(111, 42)
        Me.Button7.TabIndex = 0
        Me.Button7.Text = "Clear Line Prop."
        Me.Button7.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(105, 192)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(87, 20)
        Me.TextBox1.TabIndex = 8
        Me.TextBox1.Text = "Tag No."
        '
        'cmdFilePath
        '
        Me.cmdFilePath.Location = New System.Drawing.Point(105, 51)
        Me.cmdFilePath.Multiline = True
        Me.cmdFilePath.Name = "cmdFilePath"
        Me.cmdFilePath.Size = New System.Drawing.Size(255, 23)
        Me.cmdFilePath.TabIndex = 9
        '
        'lblFilePath
        '
        Me.lblFilePath.AutoSize = True
        Me.lblFilePath.Location = New System.Drawing.Point(103, 35)
        Me.lblFilePath.Name = "lblFilePath"
        Me.lblFilePath.Size = New System.Drawing.Size(35, 13)
        Me.lblFilePath.TabIndex = 10
        Me.lblFilePath.Text = "Path: "
        '
        'lblTestPath
        '
        Me.lblTestPath.AutoSize = True
        Me.lblTestPath.Location = New System.Drawing.Point(103, 81)
        Me.lblTestPath.Name = "lblTestPath"
        Me.lblTestPath.Size = New System.Drawing.Size(32, 13)
        Me.lblTestPath.TabIndex = 10
        Me.lblTestPath.Text = "Path:"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(106, 96)
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(255, 23)
        Me.TextBox2.TabIndex = 9
        '
        'cmdBack
        '
        Me.cmdBack.Location = New System.Drawing.Point(3, 5)
        Me.cmdBack.Name = "cmdBack"
        Me.cmdBack.Size = New System.Drawing.Size(24, 23)
        Me.cmdBack.TabIndex = 11
        Me.cmdBack.Text = "<"
        Me.cmdBack.UseVisualStyleBackColor = True
        '
        'AutoCitect
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(549, 287)
        Me.Controls.Add(Me.cmdBack)
        Me.Controls.Add(Me.lblTestPath)
        Me.Controls.Add(Me.lblFilePath)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.cmdFilePath)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.IOSetup)
        Me.Controls.Add(Me.CheckBox3)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.NumericUpDown1)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.cmdTestExcel)
        Me.Controls.Add(Me.cmdLoadFile)
        Me.Name = "AutoCitect"
        Me.Text = "AutoCitect"
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmdLoadFile As Button
    Friend WithEvents cmdTestExcel As Button
    Friend WithEvents NumericUpDown1 As NumericUpDown
    Friend WithEvents Label1 As Label
    Friend WithEvents Button3 As Button
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents CheckBox3 As CheckBox
    Friend WithEvents IOSetup As Label
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents cmdFilePath As TextBox
    Friend WithEvents lblFilePath As Label
    Friend WithEvents lblTestPath As Label
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents cmdBack As Button
End Class
