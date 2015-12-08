<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class tgt_license
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
 'NOTE: The following procedure is required by the Windows Form Designer
 'It can be modified using the Windows Form Designer.
 'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(tgt_license))
Me.logo_command = New System.Windows.Forms.Button
Me.license_pic = New System.Windows.Forms.Panel
Me.license_label = New System.Windows.Forms.Label
Me.license_pic.SuspendLayout()
Me.SuspendLayout()
'
'logo_command
'
Me.logo_command.BackColor = System.Drawing.SystemColors.Control
Me.logo_command.Cursor = System.Windows.Forms.Cursors.Default
Me.logo_command.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.logo_command.ForeColor = System.Drawing.SystemColors.ControlText
Me.logo_command.Location = New System.Drawing.Point(624, 8)
Me.logo_command.Name = "logo_command"
Me.logo_command.RightToLeft = System.Windows.Forms.RightToLeft.No
Me.logo_command.Size = New System.Drawing.Size(81, 33)
Me.logo_command.TabIndex = 1
Me.logo_command.Text = "Close"
Me.logo_command.UseCompatibleTextRendering = True
Me.logo_command.UseVisualStyleBackColor = False
'
'license_pic
'
Me.license_pic.BackColor = System.Drawing.SystemColors.Control
Me.license_pic.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.license_pic.Controls.Add(Me.license_label)
Me.license_pic.Controls.Add(Me.logo_command)
Me.license_pic.Cursor = System.Windows.Forms.Cursors.Default
Me.license_pic.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.license_pic.ForeColor = System.Drawing.SystemColors.ControlText
Me.license_pic.Location = New System.Drawing.Point(1, 0)
Me.license_pic.Name = "license_pic"
Me.license_pic.RightToLeft = System.Windows.Forms.RightToLeft.No
Me.license_pic.Size = New System.Drawing.Size(740, 640)
Me.license_pic.TabIndex = 0
Me.license_pic.TabStop = True
'
'license_label
'
Me.license_label.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.license_label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.license_label.Location = New System.Drawing.Point(55, 101)
Me.license_label.Name = "license_label"
Me.license_label.Size = New System.Drawing.Size(573, 108)
Me.license_label.TabIndex = 4
Me.license_label.Text = resources.GetString("license_label.Text")
Me.license_label.UseCompatibleTextRendering = True
'
'tgt_license
'
Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
Me.BackColor = System.Drawing.SystemColors.Control
Me.ClientSize = New System.Drawing.Size(738, 643)
Me.Controls.Add(Me.license_pic)
Me.Cursor = System.Windows.Forms.Cursors.Default
Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
Me.Location = New System.Drawing.Point(4, 23)
Me.MaximizeBox = False
Me.MinimizeBox = False
Me.Name = "tgt_license"
Me.RightToLeft = System.Windows.Forms.RightToLeft.No
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "FirstButton(tm)   License Agreement"
Me.license_pic.ResumeLayout(False)
Me.ResumeLayout(False)

End Sub
 Public WithEvents logo_command As System.Windows.Forms.Button
 Public WithEvents license_pic As System.Windows.Forms.Panel
 Friend WithEvents license_label As System.Windows.Forms.Label
#End Region
End Class