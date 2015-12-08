Option Strict Off
Option Explicit On
'-------------------------------------------------------------------------------------------'
' COPYRIGHT © 2010-2012  LIQUIDITY LIGHTHOUSE, LLC. ALL RIGHTS RESERVED.                    '
' PORTIONS ARE CONVERTED FROM CODE THAT IS COPYRIGHT © 2005-2010  LIQUIDITY LIGHTHOUSE, LLC.'
'-------------------------------------------------------------------------------------------'
'This file is part of FirstButton.                                                          '
'FirstButton is free software: you can redistribute it and/or modify it under the terms of  '
' the GNU General Public License as published by the Free Software Foundation, either       '
' version 3 of the License, or (at your option) any later version.                          '
'FirstButton is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;   '
' without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. '
' See the GNU General Public License for more details.                                      '
'You should have received a copy of the GNU General Public License along with FirstButton.  '
' If not, see <http://www.gnu.org/licenses/>.                                               '
'-------------------------------------------------------------------------------------------'
Friend Class tgt_license
	Inherits System.Windows.Forms.Form
	
	Private Sub tgt_license_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		Me.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 4), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		
	End Sub
	
	Private Sub tgt_license_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Icon = tgt_inst.Icon
		Me.Hide()
		
	End Sub
	
	Private Sub logo_command_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles logo_command.Click
		'hide the license agreement screen once the user clicks on the command button'
		Me.Hide()
		tgt_inst.Activate()
		
	End Sub

Private Sub license_pic_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles license_pic.Paint

End Sub

End Class