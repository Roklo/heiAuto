<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGenGenies
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdSelectFile = New System.Windows.Forms.Button()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.listSystems = New System.Windows.Forms.ComboBox()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.txtAddSys = New System.Windows.Forms.TextBox()
        Me.cmdGenGenies = New System.Windows.Forms.Button()
        Me.lblFilePath = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.listGenies = New System.Windows.Forms.ComboBox()
        Me.cmdBack = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(164, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(266, 25)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Generate Genies from IO-List"
        '
        'cmdSelectFile
        '
        Me.cmdSelectFile.Location = New System.Drawing.Point(79, 77)
        Me.cmdSelectFile.Name = "cmdSelectFile"
        Me.cmdSelectFile.Size = New System.Drawing.Size(90, 31)
        Me.cmdSelectFile.TabIndex = 3
        Me.cmdSelectFile.Text = "Select IO-List"
        Me.cmdSelectFile.UseVisualStyleBackColor = True
        '
        'lbl1
        '
        Me.lbl1.AutoSize = True
        Me.lbl1.Location = New System.Drawing.Point(177, 86)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(54, 13)
        Me.lbl1.TabIndex = 4
        Me.lbl1.Text = "File Path: "
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(125, 139)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(75, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Select system:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(312, 139)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(175, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "System not in list? Add system here:"
        '
        'listSystems
        '
        Me.listSystems.FormattingEnabled = True
        Me.listSystems.Location = New System.Drawing.Point(79, 155)
        Me.listSystems.Name = "listSystems"
        Me.listSystems.Size = New System.Drawing.Size(169, 21)
        Me.listSystems.TabIndex = 6
        '
        'cmdAdd
        '
        Me.cmdAdd.Location = New System.Drawing.Point(466, 156)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(35, 20)
        Me.cmdAdd.TabIndex = 7
        Me.cmdAdd.Text = "Add"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'txtAddSys
        '
        Me.txtAddSys.Location = New System.Drawing.Point(291, 156)
        Me.txtAddSys.Name = "txtAddSys"
        Me.txtAddSys.Size = New System.Drawing.Size(169, 20)
        Me.txtAddSys.TabIndex = 8
        '
        'cmdGenGenies
        '
        Me.cmdGenGenies.Location = New System.Drawing.Point(209, 300)
        Me.cmdGenGenies.Name = "cmdGenGenies"
        Me.cmdGenGenies.Size = New System.Drawing.Size(121, 47)
        Me.cmdGenGenies.TabIndex = 9
        Me.cmdGenGenies.Text = "Generate Genies"
        Me.cmdGenGenies.UseVisualStyleBackColor = True
        '
        'lblFilePath
        '
        Me.lblFilePath.Location = New System.Drawing.Point(232, 86)
        Me.lblFilePath.Name = "lblFilePath"
        Me.lblFilePath.Size = New System.Drawing.Size(269, 13)
        Me.lblFilePath.TabIndex = 4
        Me.lblFilePath.Text = "......"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(203, 200)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(131, 13)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Select genies to generate:"
        '
        'listGenies
        '
        Me.listGenies.FormattingEnabled = True
        Me.listGenies.Location = New System.Drawing.Point(184, 216)
        Me.listGenies.Name = "listGenies"
        Me.listGenies.Size = New System.Drawing.Size(169, 21)
        Me.listGenies.TabIndex = 6
        '
        'cmdBack
        '
        Me.cmdBack.Location = New System.Drawing.Point(3, 5)
        Me.cmdBack.Name = "cmdBack"
        Me.cmdBack.Size = New System.Drawing.Size(24, 23)
        Me.cmdBack.TabIndex = 12
        Me.cmdBack.Text = "<"
        Me.cmdBack.UseVisualStyleBackColor = True
        '
        'frmGenGenies
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(576, 359)
        Me.Controls.Add(Me.cmdBack)
        Me.Controls.Add(Me.cmdGenGenies)
        Me.Controls.Add(Me.txtAddSys)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.listGenies)
        Me.Controls.Add(Me.listSystems)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblFilePath)
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.cmdSelectFile)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmGenGenies"
        Me.Text = "AutoCitect v0.00001 - Generate Genies"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents cmdSelectFile As Button
    Friend WithEvents lbl1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents listSystems As ComboBox
    Friend WithEvents cmdAdd As Button
    Friend WithEvents txtAddSys As TextBox
    Friend WithEvents cmdGenGenies As Button
    Friend WithEvents lblFilePath As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents listGenies As ComboBox
    Friend WithEvents cmdBack As Button
End Class
