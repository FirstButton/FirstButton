<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class tgt_inst
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
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents title_pic As System.Windows.Forms.PictureBox
	Public WithEvents cancel_cmd As System.Windows.Forms.Button
	Public WithEvents next_cmd As System.Windows.Forms.Button
	Public WithEvents agree_check As System.Windows.Forms.CheckBox
	Public WithEvents read_agmt_cmd As System.Windows.Forms.Button
	Public WithEvents uninstall_option As System.Windows.Forms.RadioButton
	Public WithEvents install_option As System.Windows.Forms.RadioButton
	Public WithEvents sample_lbl As System.Windows.Forms.Label
	Public WithEvents install_frame As System.Windows.Forms.GroupBox
	Public WithEvents cpyright_lbl As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(tgt_inst))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.title_pic = New System.Windows.Forms.PictureBox()
        Me.cancel_cmd = New System.Windows.Forms.Button()
        Me.next_cmd = New System.Windows.Forms.Button()
        Me.agree_check = New System.Windows.Forms.CheckBox()
        Me.read_agmt_cmd = New System.Windows.Forms.Button()
        Me.install_frame = New System.Windows.Forms.GroupBox()
        Me.uninstall_option = New System.Windows.Forms.RadioButton()
        Me.install_option = New System.Windows.Forms.RadioButton()
        Me.sample_lbl = New System.Windows.Forms.Label()
        Me.cpyright_lbl = New System.Windows.Forms.Label()
        Me.IE6_label = New System.Windows.Forms.Label()
        CType(Me.title_pic, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.install_frame.SuspendLayout()
        Me.SuspendLayout()
        '
        'title_pic
        '
        Me.title_pic.BackColor = System.Drawing.Color.White
        Me.title_pic.Cursor = System.Windows.Forms.Cursors.Default
        Me.title_pic.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.title_pic.ForeColor = System.Drawing.SystemColors.ControlText
        Me.title_pic.Image = CType(resources.GetObject("title_pic.Image"), System.Drawing.Image)
        Me.title_pic.Location = New System.Drawing.Point(2, 8)
        Me.title_pic.Name = "title_pic"
        Me.title_pic.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.title_pic.Size = New System.Drawing.Size(193, 105)
        Me.title_pic.TabIndex = 8
        Me.title_pic.TabStop = False
        '
        'cancel_cmd
        '
        Me.cancel_cmd.BackColor = System.Drawing.Color.White
        Me.cancel_cmd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cancel_cmd.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cancel_cmd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cancel_cmd.Location = New System.Drawing.Point(258, 358)
        Me.cancel_cmd.Name = "cancel_cmd"
        Me.cancel_cmd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cancel_cmd.Size = New System.Drawing.Size(97, 31)
        Me.cancel_cmd.TabIndex = 6
        Me.cancel_cmd.Text = "Cancel"
        Me.cancel_cmd.UseCompatibleTextRendering = True
        Me.cancel_cmd.UseVisualStyleBackColor = False
        '
        'next_cmd
        '
        Me.next_cmd.BackColor = System.Drawing.Color.White
        Me.next_cmd.Cursor = System.Windows.Forms.Cursors.Default
        Me.next_cmd.Font = New System.Drawing.Font("Arial", 11.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.next_cmd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.next_cmd.Location = New System.Drawing.Point(98, 356)
        Me.next_cmd.Name = "next_cmd"
        Me.next_cmd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.next_cmd.Size = New System.Drawing.Size(113, 33)
        Me.next_cmd.TabIndex = 5
        Me.next_cmd.Text = "Next"
        Me.next_cmd.UseCompatibleTextRendering = True
        Me.next_cmd.UseVisualStyleBackColor = False
        '
        'agree_check
        '
        Me.agree_check.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.agree_check.Cursor = System.Windows.Forms.Cursors.Default
        Me.agree_check.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.agree_check.ForeColor = System.Drawing.SystemColors.ControlText
        Me.agree_check.Location = New System.Drawing.Point(98, 120)
        Me.agree_check.Name = "agree_check"
        Me.agree_check.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.agree_check.Size = New System.Drawing.Size(281, 41)
        Me.agree_check.TabIndex = 4
        Me.agree_check.Text = "I agree to the terms of the License Agreement"
        Me.agree_check.UseCompatibleTextRendering = True
        Me.agree_check.UseVisualStyleBackColor = False
        '
        'read_agmt_cmd
        '
        Me.read_agmt_cmd.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.read_agmt_cmd.Cursor = System.Windows.Forms.Cursors.Default
        Me.read_agmt_cmd.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.read_agmt_cmd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.read_agmt_cmd.Location = New System.Drawing.Point(245, 51)
        Me.read_agmt_cmd.Name = "read_agmt_cmd"
        Me.read_agmt_cmd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.read_agmt_cmd.Size = New System.Drawing.Size(161, 21)
        Me.read_agmt_cmd.TabIndex = 3
        Me.read_agmt_cmd.Text = "Read License Agreement"
        Me.read_agmt_cmd.UseCompatibleTextRendering = True
        Me.read_agmt_cmd.UseVisualStyleBackColor = False
        '
        'install_frame
        '
        Me.install_frame.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.install_frame.Controls.Add(Me.uninstall_option)
        Me.install_frame.Controls.Add(Me.install_option)
        Me.install_frame.Controls.Add(Me.sample_lbl)
        Me.install_frame.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.install_frame.ForeColor = System.Drawing.Color.White
        Me.install_frame.Location = New System.Drawing.Point(21, 200)
        Me.install_frame.Name = "install_frame"
        Me.install_frame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.install_frame.Size = New System.Drawing.Size(416, 129)
        Me.install_frame.TabIndex = 0
        Me.install_frame.TabStop = False
        Me.install_frame.UseCompatibleTextRendering = True
        '
        'uninstall_option
        '
        Me.uninstall_option.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.uninstall_option.Cursor = System.Windows.Forms.Cursors.Default
        Me.uninstall_option.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.uninstall_option.ForeColor = System.Drawing.SystemColors.ControlText
        Me.uninstall_option.Location = New System.Drawing.Point(32, 88)
        Me.uninstall_option.Name = "uninstall_option"
        Me.uninstall_option.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.uninstall_option.Size = New System.Drawing.Size(297, 33)
        Me.uninstall_option.TabIndex = 2
        Me.uninstall_option.TabStop = True
        Me.uninstall_option.Text = "Uninstall Toolbar Button "
        Me.uninstall_option.UseCompatibleTextRendering = True
        Me.uninstall_option.UseVisualStyleBackColor = False
        '
        'install_option
        '
        Me.install_option.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.install_option.Checked = True
        Me.install_option.Cursor = System.Windows.Forms.Cursors.Default
        Me.install_option.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.install_option.ForeColor = System.Drawing.SystemColors.ControlText
        Me.install_option.Location = New System.Drawing.Point(32, 40)
        Me.install_option.Name = "install_option"
        Me.install_option.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.install_option.Size = New System.Drawing.Size(336, 41)
        Me.install_option.TabIndex = 1
        Me.install_option.TabStop = True
        Me.install_option.Text = "Install Toolbar Button for Internet Explorer"
        Me.install_option.UseCompatibleTextRendering = True
        Me.install_option.UseVisualStyleBackColor = False
        '
        'sample_lbl
        '
        Me.sample_lbl.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.sample_lbl.Cursor = System.Windows.Forms.Cursors.Default
        Me.sample_lbl.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sample_lbl.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.sample_lbl.Location = New System.Drawing.Point(160, 16)
        Me.sample_lbl.Name = "sample_lbl"
        Me.sample_lbl.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.sample_lbl.Size = New System.Drawing.Size(128, 22)
        Me.sample_lbl.TabIndex = 9
        Me.sample_lbl.Text = "Sample Version"
        Me.sample_lbl.UseCompatibleTextRendering = True
        '
        'cpyright_lbl
        '
        Me.cpyright_lbl.BackColor = System.Drawing.Color.Transparent
        Me.cpyright_lbl.Cursor = System.Windows.Forms.Cursors.Default
        Me.cpyright_lbl.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cpyright_lbl.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cpyright_lbl.Location = New System.Drawing.Point(100, 416)
        Me.cpyright_lbl.Name = "cpyright_lbl"
        Me.cpyright_lbl.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cpyright_lbl.Size = New System.Drawing.Size(263, 17)
        Me.cpyright_lbl.TabIndex = 7
        Me.cpyright_lbl.Text = "(c) 2005-2015 Copyright Liquidity Lighthouse, LLC."
        Me.cpyright_lbl.UseCompatibleTextRendering = True
        '
        'IE6_label
        '
        Me.IE6_label.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IE6_label.ForeColor = System.Drawing.Color.Turquoise
        Me.IE6_label.Location = New System.Drawing.Point(72, 165)
        Me.IE6_label.Name = "IE6_label"
        Me.IE6_label.Size = New System.Drawing.Size(321, 21)
        Me.IE6_label.TabIndex = 9
        Me.IE6_label.Text = "IE6 Users - Save and run outside of Internet Explorer"
        Me.IE6_label.UseCompatibleTextRendering = True
        '
        'tgt_inst
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(458, 436)
        Me.Controls.Add(Me.IE6_label)
        Me.Controls.Add(Me.title_pic)
        Me.Controls.Add(Me.cancel_cmd)
        Me.Controls.Add(Me.next_cmd)
        Me.Controls.Add(Me.agree_check)
        Me.Controls.Add(Me.read_agmt_cmd)
        Me.Controls.Add(Me.install_frame)
        Me.Controls.Add(Me.cpyright_lbl)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.White
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(4, 23)
        Me.MaximizeBox = False
        Me.Name = "tgt_inst"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FirstButton(tm)  Installer"
        CType(Me.title_pic, System.ComponentModel.ISupportInitialize).EndInit()
        Me.install_frame.ResumeLayout(False)
        Me.ResumeLayout(False)

End Sub
 Friend WithEvents IE6_label As System.Windows.Forms.Label
#End Region 
End Class