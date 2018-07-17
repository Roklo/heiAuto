<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMenu
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.cmdTestPage = New System.Windows.Forms.Button()
        Me.cmdGenGenies = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(122, 106)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(113, 59)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Menu Option 1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(122, 171)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(113, 59)
        Me.Button2.TabIndex = 0
        Me.Button2.Text = "Menu Option 2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(122, 236)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(113, 59)
        Me.Button3.TabIndex = 0
        Me.Button3.Text = "Menu Option 3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'cmdTestPage
        '
        Me.cmdTestPage.Location = New System.Drawing.Point(293, 236)
        Me.cmdTestPage.Name = "cmdTestPage"
        Me.cmdTestPage.Size = New System.Drawing.Size(113, 59)
        Me.cmdTestPage.TabIndex = 0
        Me.cmdTestPage.Text = "Test Page"
        Me.cmdTestPage.UseVisualStyleBackColor = True
        '
        'cmdGenGenies
        '
        Me.cmdGenGenies.Location = New System.Drawing.Point(293, 171)
        Me.cmdGenGenies.Name = "cmdGenGenies"
        Me.cmdGenGenies.Size = New System.Drawing.Size(113, 59)
        Me.cmdGenGenies.TabIndex = 0
        Me.cmdGenGenies.Text = "Generate Genies From IO-List"
        Me.cmdGenGenies.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(293, 106)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(113, 59)
        Me.Button6.TabIndex = 0
        Me.Button6.Text = "Menu Option 4"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(172, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(189, 25)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "AutoCitect v0.00001"
        '
        'frmMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(540, 324)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.cmdGenGenies)
        Me.Controls.Add(Me.cmdTestPage)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "frmMenu"
        Me.Text = "AutoCitect v0.00001 - Menu"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents cmdTestPage As Button
    Friend WithEvents cmdGenGenies As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Label1 As Label
End Class
