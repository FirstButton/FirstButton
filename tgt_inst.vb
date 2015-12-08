Option Strict Off
Option Explicit On
Imports Microsoft.Win32
Imports VB = Microsoft.VisualBasic
Imports System.Collections

Friend Class tgt_inst
    Inherits System.Windows.Forms.Form
    Public Class myReverserClass
        Implements IComparer
        ' Calls CaseInsensitiveComparer.Compare with the parameters reversed.
        Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
           Implements IComparer.Compare
            Return New CaseInsensitiveComparer().Compare(x, y)
        End Function
    End Class
'-------------------------------------------------------------------------------------------'
' THIS PROGRAM IS CONFIDENTIAL AND PROPRIETARY TO LIQUIDITY LIGHTHOUSE, LLC                 '
'  AND MAY NOT BE REPRODUCED, PUBLISHED, OR DISCLOSED TO OTHERS WITHOUT                     '
'   COMPANY AUTHORIZATION.                                                                  '
' COPYRIGHT © 2010-2015  LIQUIDITY LIGHTHOUSE, LLC. ALL RIGHTS RESERVED.                    '
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
Dim logo_cmdmap_spread_above As Integer    '01/24/2011'
Dim logo_cmdmap_spread_below As Integer    '01/24/2011'
Dim logo_current_nextid_value As Integer   '01/24/2011'
Dim logo_current_nextid_inuse As Boolean   '01/24/2011'
Dim logo_nextid_spread_changed As Boolean  '01/24/2011'
Dim edit_string As String
Dim edit_length As Integer
Dim hKey As Integer
Dim logo_high_dblword_value As Integer
Dim rk As RegistryKey
Dim create_rk As RegistryKey
Dim rk2 As RegistryKey
Dim rk3 As RegistryKey
Dim logo_cmdvalue_obj As Object
Dim svaluename As String
Dim skeyname As String
Public logo_install_directory As String
Dim logo_new_guid As String
Dim logo_new_guid_2 As String
Dim logo_brand_name As String
Dim logo_company_name As String
Dim logo_company_url As String
Dim logo_install_module As String
Dim logo_install_module2 As String
Dim logo_install_module3 As String
Dim logo_ie_ext_flags_value As String
Dim logo_current_ext_status As String
Dim logo_nextid_missing_flag As Boolean
Dim ie7_value_returned As Object
Dim ie6_value_returned As Object
Dim ie7_value_returned_str As String
Dim logo_install_guid As String
Dim logo_install_guid_2 As String
Dim logo_IE7_install_guid As String
Dim logo_IE7_install_guid_2 As String
Dim logo_toolbar_cmdmap_exists As String
Dim logo_IE7_toolbar_cmdmap_exists As String
Dim logo_b4_toolbar_cmdmap_exists As String
Dim IE7_flag As Boolean
Dim IE8_flag As Boolean
Dim IE9_flag As Boolean
Dim IE10_flag As Boolean
Dim IE11_flag As Boolean                        '11/26/2013'
Dim create_prob_flag As Boolean
Dim cmd_id_compare_value As Integer
Dim cmd_id_add_value As Integer
Dim logo_IE7_cmdid_installed As Boolean
Dim logo_b4_cmdid_installed As Boolean
Dim logo_IE7_toolbar_prev_install As Boolean
Dim logo_b4_toolbar_prev_install As Boolean
Dim logo_b4_proc_flag As Boolean
Dim logo_other_ie7_cmds_flag As Boolean
Dim logo_IE7_bumped_cmdmap_flag As Boolean
Dim logo_IE7_build_cmdmap_values As Boolean
Dim logo_cmdmapping_flag As Boolean
Dim logo_b4_cmdmapping_flag As Boolean
Dim logo_default_toolbar_built As Boolean
Dim logo_IE7_found_ext_key As Boolean
Dim logo_IE7_default_toolbar_built As Boolean
Dim reg_ext_error As Boolean
Dim logo_toolbar_prev_install As Boolean
Dim logo_expand_path_flag As Boolean
Dim logo_install_activity As Boolean
Dim logo_auth_copy_system As Boolean
Dim logo_auth_copy_programs As Boolean
Dim logo_source_not_found As Boolean
Dim logo_allow_multi_installs As Boolean
Dim logo_IE7_allow_multi_installs As Boolean
Dim logo_install_favorite As Boolean
Dim logo_install_tools As Boolean
Dim logo_install_add_remove As Boolean
Dim programs_not_found As Boolean
Dim logo_other_error_flag As Boolean
Dim logo_auth_copy_program As Boolean
Dim logo_install_index As Integer
Dim logo_install_index_2 As Integer
Dim logo_IE7_install_index As Integer
Dim logo_IE7_install_index_2 As Integer
Dim logo_next_avail_id As Integer
Dim IE7_logo_next_avail_id As Integer
Dim x As Integer
Dim logo_guid_map_num As Integer
Dim logo_guid_map_num_2 As Integer
Dim logo_guid_map_byte As Byte
Dim logo_guid_map_byte_2 As Byte
Dim h21_hi_nib As Integer
Dim h21_lo_nib As Integer
Dim h21_combined As Integer
Dim h21_eval_char As String
Dim h31_hi_nib As Integer
Dim h31_lo_nib As Integer
Dim shtl_eval_char As String
Dim shtl_21_long As Integer
Dim shtl_lo As Integer
Dim shtl_hi As Integer
Dim logo_num_bytes_to_move As Integer
Dim logo_IE7_num_bytes_to_move As Integer
Dim logo_IE7_start_byte_pos As Integer
Dim logo_IE7_button_difference As Integer
Dim num_of_IE7_logo_buttons As Integer
Dim num_of_b4_logo_buttons As Integer
Dim logo_start_byte_pos As Integer
Dim logo_button_difference As Integer
Dim num_of_logo_buttons As Integer
Dim logo_guid_map_string As String
Dim logo_check_parm_name As String
Dim logo_check_counter As Integer
Dim filename_counter As Integer
Public logo_error_code As Integer
Public logo_error_source As String
Public logo_error_action As String
Public logo_error_text As String
Dim logo_lreturnvalue As Long
Dim logo_return_value As Long
Dim logo_create_rc As Long
Dim logo_IE7_lreturnvalue As Integer
Dim logo_fav_retvalue As Long
Dim logo_return_code As Integer
Dim netnewpath As String
Dim logo_exe_path As String
Dim logo_i As Integer
Dim logo_d As Integer
Dim logo_e As Integer
Dim logo_IE7_i As Integer
Dim logo_IE7_j As Integer
Dim IE7_version_string As String
Dim logo_IE7_d As Integer
Dim logo_IE7_e As Integer
Dim match_b4_loop As Integer
Dim match_ie_loop As Integer
Dim logo_b4_i As Integer
Dim cmdmap_loop As Integer
Dim logo_favorite_path_len As Integer
Dim logo_source_path As String
Dim logo_app_path As String
Dim logo_favorite_dir As String
Dim logo_favorite_path As String
Dim logo_system_dir As String
Dim logo_programs_dir As String
Dim logo_programs_path_len As Integer
Dim logo_systemroot_not_found As Boolean
Dim logo_source_file As String
Dim logo_source_dir As String
Dim logo_dest_file As String
Dim logo_long_returned As Integer
Dim logo_long_byte As Object
Dim logo_long_ext_byte As Object
Dim logo_long_binary As Byte
Dim logo_string_returned As String
Dim logo_string_ret_length As Integer
Dim logo_num_of_toolbar_items As Integer
Dim logo_b4_string_ret_length As Integer
Dim logo_IE7_string_ret_length As Integer
Dim logo_IE7_toolbar_remain As Integer
Dim logo_b4_num_of_toolbar_items As Integer
Dim logo_IE7_num_of_toolbar_items As Integer
Dim logo_filename_counter As Integer
Dim logo_string_filename As String
Dim logo_begin_button_num As Integer
Dim logo_IE7_begin_button_num As Integer
Dim logo_b4_begin_button_num As Integer
Dim logo_IE7_default_separator As Integer
Dim logo_first_separator As Integer
Dim logo_default_separator As Integer
Dim logo_first_matched_id As String
Dim logo_first_mapid_pos As Integer
Dim logo_mapid_found_flag As Boolean
Dim logo_start_pos As Integer
Dim logo_IE7_first_mapid_pos As Integer
Dim logo_IE7_first_mapid_found_flag As Boolean
Dim logo_IE7_start_pos As Integer
Dim logo_IE7_target_button_pos As Integer
Dim logo_IE7_separator_found As Boolean
Dim logo_target_button_pos As Integer
Dim logo_separator_found As Boolean
Dim logo_num_of_separators As Integer
Dim logo_num_of_IE7_separators As Integer
Dim logo_pos_sep_1 As Integer
Dim logo_pos_sep_2 As Integer
Dim logo_IE7_pos_sep_1 As Integer
Dim logo_IE7_pos_sep_2 As Integer
Dim logo_default_target_pos As Integer
Dim logo_button_installed As Boolean
Dim logo_current_button_num As Integer
Dim logo_last_button_seq As Integer
Dim logo_IE7_button_installed As Boolean
Dim logo_IE7_default_target_pos As Integer
Dim logo_IE7_target_compare_pos As Integer
Dim logo_IE7_current_button_num As Integer
Dim logo_IE7_last_button_seq As Integer
Dim logo_b4_last_button_seq As Integer
Dim logo_last_button_value As Long
Dim logo_last_button_value_2 As Long
Dim logo_seq_button_value As Integer
Dim logo_seq_button_value_2 As Integer
Dim logo_default_toolbar_map() As Byte
Dim logo_default_toolbar_length As Integer
Dim logo_default_toolbar_byte() As Byte
Dim logo_IE7_last_button_value As Integer
Dim logo_b4_last_button_value As Integer
Dim logo_IE7_last_button_value_2 As Integer
Dim logo_IE7_seq_button_value As Integer
Dim logo_IE7_seq_button_value_2 As Integer
Dim logo_IE7_default_toolbar_map() As Byte
Dim logo_default_IE7_toolbar_length As Integer
Dim logo_IE7_default_toolbar_byte() As Byte
Dim logo_b4_IE7_default_toolbar_byte() As Byte
Dim logo_IE7_default_toolbar_length As Integer
Dim logo_default_toolbar_var As Object
Dim dest_lo_value As Integer
Dim dest_hi_value As Integer
Dim dest_lo_string As String
Dim dest_hi_string As String
Dim dest_string As String
Dim dest_lo_int As Short
Dim dest_hi_int As Short
Dim toolbar_string As String
Public logo_sub_keyname As String
Dim logo_subkey As String
Dim logo_IE7_subkey As String
Dim logo_IE7_subkeyname As String
Dim logo_IE7_sub_valuename As String
Dim logo_string_key_value As String
Public logo_sub_valuename As String
Public logo_uni_valuename As String
Dim logo_filename_one_ As String
Dim logo_filename_two As String
Dim logo_file_number As Integer
Dim logo_file_status As Integer
Dim logo_file_total_bytes As Decimal
Dim logo_file_free_bytes As Decimal
Dim logo_file_free_bytes_avail As Decimal
Dim conv_free_space As Decimal
Dim enough_space_flag As String
Dim logo_hive_registry_key As Integer
Dim logo_hive_cmdmap_key As Integer
Dim logo_b4_registry_key As Integer
Dim logo_IE7_string_key_value As String
Dim logo_file_s_ As String
Dim dl As Integer
Dim logo_hive_key As Integer
Dim filedefn As String
Dim filechar As String
Dim logo_file_character As String
Dim source_length As Integer
Dim logo_byte_string() As Byte
Dim logo_IE7_byte_string() As Byte
Dim logo_b4_byte_string() As Byte
Public win2kxp_flag As String
Public win2000_flag As String
Public win_vista_flag As String     '11/21/2013'
Public win7_flag As String          '11/21/2013'
Public win8_flag As String          '04/21/2012'
Public winverflag As String
Dim win_UAC_enabled_flag As Boolean          '11/10/2013'
Dim logo_check_UAC_flag As Boolean            '11/10/2013'
Dim logo_rqve_hkey As Integer
Dim logo_rqve_lpclass As String
Dim logo_rqve_lpcbclass As Integer
Dim logo_rqve_lpcsubkeys As Integer
Dim logo_rqve_lpcbmaxsubkeylen As Integer
Dim logo_rqve_lpcbmaxclasslen As Integer
Dim logo_rqve_lpcvalues As Integer
Dim logo_rqve_lpcbmaxvaluenamelen As Integer
Dim logo_rqve_lpcbmaxvaluelen As Integer
Dim logo_rqve_lpcbsecuritydesc As Integer
Dim rev_hkey As Integer
Dim rev_dwindex As Integer
Dim rev_lpvaluename As String
Dim rev_lpcbvaluename As Integer
Dim rev_lptype As Integer
Dim rev_lpdata As Byte
Dim rev_lpcbdata As Integer
Dim cmd_hkey As Integer
Dim key_hkey As Integer
Dim key_dwindex As Integer
Dim key_strvalue As String
Dim key_strvaluelen As Integer 
Dim key_strclass As String
Dim key_strclasslen As Integer
Dim key_strresult As String 
Dim cmd_ext_array(90) As extension_array
Dim cmdmapping_array(90) As cmdmapping_entry
Dim cmdarray(90) As cmdmap_entry
Dim cmdarray_sort(90) As cmdmap_entry
Dim cmd_IE7_array(90) As cmdmap_entry
Dim cmd_IE7_array_sort_2(90) As cmdmap_entry
Dim cmd_IE7_array_hklm(90) As cmdmap_entry
Dim cmd_IE7_b4_array_hklm(90) As ie7_cmdmap_entry
Dim cmd_IE7_array_sort(90) As ie7_cmdmap_entry
Dim cmd_IE7_b4_array(90) As ie7_cmdmap_entry
Dim guid_sort_array() As String
Dim cmdmap_sort_array() As String
Dim cmdmapping_sort_array() As String
Dim IE7_search_guid As String
Dim IE7_search_length As Integer
Dim IE7_cmdcount As Integer
Dim IE7_b4_cmdcount As Integer
Dim IE7_cmdcount_hklm As Integer
Dim IE7_b4_cmdcount_hklm As Integer
Dim IE7_b4_cmdcount_hklm_loop As Integer
Dim IE7_b4_cmdcount_hklm_both As Integer
Dim IE7_cmdcount_hklm_loop As Integer
Dim IE7_cmdcount_hklm_both As Integer
Dim logo_found_in_hive As String
Dim IE7_cmd_ctr As Integer
Dim IE7_current_cmd_id As Integer
Dim b4_current_cmd_id As Integer 
Dim IE7_search_cmd_id As Integer 
Dim current_cmdmap_id As Integer 
Dim toolbar_IE7_cmd_count As Integer
Dim logo_IE7_use_mapid_logic As Boolean
Dim max_poss_cmd_id As Integer
Dim cmdcount As Integer
Dim cmdcount_2 As Integer 
Dim logo_num_cmdmaps As Integer 
Dim mapcount As Integer     
Dim toolbar_cmd_array(60) As toolbar_map_entry
Dim toolbar_cmd_count As Integer
Dim cmdmap_diff As Integer
Dim toolbar_cmd_string As String
Dim deinst_toolbar_flag As Boolean
Dim logo_use_mapid_logic As Boolean
Dim logo_run_from_tempdir As Boolean
Dim prev_instance_flag As Boolean
Dim logo_IE_final_msg As String        '02/17/2013'
Public Sub display_logo_error_msg()
    Dim style As Object
    style = MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation
    MsgBox(Me.logo_error_text, MsgBoxStyle.Exclamation, "FirstButton  Installer")
End Sub
Private Sub success_install_msg()
  install_ext_settings()
  If create_prob_flag = True Then
   installation_error()
   cancel_button_Click()                    '02/17/2013'
   Exit Sub                                 '02/17/2013'
  End If
  If logo_button_installed = True Or logo_IE7_button_installed = True Then
      x = MsgBox("Successfully Installed. " + logo_IE_final_msg, MsgBoxStyle.Information, "FirstButton  Installer")
      cancel_button_Click()
      Exit Sub                             '02/17/2013'
  Else
      x = MsgBox("Successfully Installed-02. " + logo_IE_final_msg, MsgBoxStyle.Information, "FirstButton  Installer")
      cancel_button_Click()                '02/17/2013'
      Exit Sub                             '02/17/2013'
  End If
End Sub
Private Sub success_delete_msg()
 x = MsgBox("Uninstall Successful. " + logo_IE_final_msg, MsgBoxStyle.Information, "FirstButton  Installer")
  cancel_button_Click()
  Exit Sub                                 '02/17/2013'
End Sub
Private Sub cancel_button_Click()  
  Me.Close()
  Exit Sub                                 '02/17/2013'
End Sub
Private Sub cancel_cmd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cancel_cmd.Click
  cancel_button_Click()
  Exit Sub                                 '02/17/2013'
End Sub 
Private Sub tgt_inst_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
       Me.WindowState = System.Windows.Forms.FormWindowState.Normal
    install_option.Checked = True
    uninstall_option.Checked = False
    install_option.TabStop = False
    uninstall_option.TabStop = False
    IE7_flag = False
    IE8_flag = False                      '11/25/2013'
    IE9_flag = False                      '11/25/2013'
    IE10_flag = False                     '11/25/2013'
    logo_other_ie7_cmds_flag = False
    logo_IE7_default_toolbar_built = False
    logo_IE7_toolbar_prev_install = False
    logo_b4_proc_flag = True
    logo_IE7_build_cmdmap_values = False    
    logo_run_from_tempdir = False
    create_prob_flag = False
    win_UAC_enabled_flag = False          '11/10/2013'
    logo_check_UAC_flag = False           '11/10/2013'
    getversion_somb()
    get_IE_version()

 '   If logo_check_UAC_flag = True Then    '11/10/2013'
    'check registry to see if UAC is enabled'
 '     check_UAC_setting()                 '11/10/2013'
 '   End If
 '   If win_UAC_enabled_flag = True Then   '11/10/2013'
      'check to see if this program was launched with Run as Administrator'
 '     MsgBox("User Account Control is enabled on this computer. If this program was not run with Administor authority, it may fail." + vbCr + vbCr + "If error occurs - please Run As Administrator, outside of Internet Explorer.", _
 '      MsgBoxStyle.Information, "FirstButton  Installer")          '11/10/2013'
 '   End If                                '11/10/2013'

    logo_source_path = My.Application.Info.DirectoryPath
    If Len(logo_source_path) >= 6 Then
      For logo_i = 1 To Len(logo_source_path)
          If Mid(logo_source_path, logo_i, 6) = "Tempor" Or Mid(logo_source_path, logo_i, 6) = "tempor" Then
              logo_run_from_tempdir = True
          End If
      Next
    End If
    If VB.Right(logo_source_path, 1) = "\" Then
        logo_source_path = My.Application.Info.DirectoryPath
    Else
        logo_source_path = My.Application.Info.DirectoryPath & "\"
    End If

    logo_company_name = "Sombreros in the Sun LLC"
    logo_brand_name = "Sombreros in the Sun"
    logo_company_url = "www.sombrerosinthesun.com"
    logo_new_guid = "{DED51647-3FFD-48da-9B48-814A053F6388}"
    logo_new_guid_2 = "{D294CEA5-4017-4cb1-8B95-054C4A2FB2AE}"
    logo_install_module = "sombreros.exe"
    logo_install_module2 = "sombreros[1].exe"
    logo_install_module3 = "sombreros[2].exe"
    logo_IE7_install_guid = "N"
    logo_IE7_install_guid_2 = "N"
    logo_install_guid = "N"
    logo_install_guid_2 = "N"
    logo_toolbar_cmdmap_exists = "N"
    logo_button_installed = False
    logo_default_toolbar_built = False
    logo_toolbar_prev_install = False
    logo_default_target_pos = 2
    logo_IE7_default_target_pos = 2
    logo_IE7_use_mapid_logic = True
    logo_IE7_allow_multi_installs = False
    logo_use_mapid_logic = True
    logo_allow_multi_installs = False
    logo_install_tools = True
    logo_install_favorite = True
    logo_install_add_remove = True

End Sub
Private Function hex2lng(ByRef hexinput As String) As Integer
    If Len(hexinput) > 1 Then
        h21_eval_char = Mid(hexinput, 1, 1)
        If IsNumeric(h21_eval_char) Then
            h21_hi_nib = 16 * CInt(Mid(hexinput, 1, 1))
        Else
            Select Case h21_eval_char
                Case "A", "a"
                    h21_hi_nib = 160
                Case "B", "b"
                    h21_hi_nib = 176
                Case "C", "c"
                    h21_hi_nib = 192
                Case "D", "d"
                    h21_hi_nib = 208
                Case "E", "e"
                    h21_hi_nib = 224
                Case "F", "f"
                    h21_hi_nib = 240
            End Select
        End If
        h21_eval_char = Mid(hexinput, 2, 1)
        If IsNumeric(h21_eval_char) Then
            h21_lo_nib = CInt(Mid(hexinput, 2, 1))
        Else
            Select Case h21_eval_char
                Case "A", "a"
                    h21_lo_nib = 10
                Case "B", "b"
                    h21_lo_nib = 11
                Case "C", "c"
                    h21_lo_nib = 12
                Case "D", "d"
                    h21_lo_nib = CInt(13)
                Case "E", "e"
                    h21_lo_nib = 14
                Case "F", "f"
                    h21_lo_nib = 15
                Case Else
                    h21_lo_nib = CInt(Mid(hexinput, 2, 1))
            End Select
        End If
    End If
    h21_combined = h21_hi_nib
    h21_combined = h21_combined + h21_lo_nib
    hex2lng = h21_combined
End Function
Private Function hex2lng2(ByRef hexinput2 As String) As Integer
    If Len(hexinput2) > 1 Then
        h31_hi_nib = 4096 * CInt(Mid(hexinput2, 1, 1))
        h31_lo_nib = 256 * CInt(Mid(hexinput2, 2, 1))
    End If
    hex2lng2 = h31_hi_nib + h31_lo_nib
End Function
Private Function find_next_avail_id() As Integer
  Dim cmdmap_loop As Integer
  Dim cmdmap_loop_2 As Integer
  Dim all_values_invalid As Boolean
  cmdmap_loop_2 = 0
  If cmdcount > 0 Then
      all_values_invalid = True
      For cmdmap_loop = 0 To cmdcount - 1
          If cmdarray(cmdmap_loop).ce_valuedword > 8192 And _
             cmdarray(cmdmap_loop).ce_valuedword < 8702 Then
              all_values_invalid = False
          Else
              '  '                                           
          End If
      Next
      If all_values_invalid = True Then
          logo_next_avail_id = 8192
          Exit Function
      End If
      For cmdmap_loop = 0 To cmdcount - 1
          cmdmap_loop_2 = cmdmap_loop + 1
          If (cmdarray(cmdmap_loop).ce_valuedword >= 8192 And _
          cmdarray(cmdmap_loop_2).ce_valuedword >= 8192) And _
          (cmdarray(cmdmap_loop).ce_valuedword <= 8701 And _
          cmdarray(cmdmap_loop_2).ce_valuedword <= 8701) And _
          (cmdarray(cmdmap_loop).ce_valuedword < _
          cmdarray(cmdmap_loop_2).ce_valuedword) Then
              cmdmap_diff = cmdarray(cmdmap_loop_2).ce_valuedword - cmdarray(cmdmap_loop).ce_valuedword
              If cmdmap_diff >= 4 Then
                  logo_next_avail_id = (cmdarray(cmdmap_loop).ce_valuedword)
                  logo_next_avail_id = logo_next_avail_id + 1
                  Exit Function
              End If
          Else
            If cmdarray(cmdmap_loop).ce_valuedword >= 8192 And _
               cmdarray(cmdmap_loop).ce_valuedword <= 8699 And _
               cmdarray(cmdmap_loop_2).ce_valuedword > 8699 Then
               logo_next_avail_id = (cmdarray(cmdmap_loop).ce_valuedword)
               logo_next_avail_id = logo_next_avail_id + 1
               Exit Function
            End If
          End If
      Next
      cmdmap_diff = (8699 - cmdarray(cmdcount - 1).ce_valuedword)
      If cmdmap_diff >= 4 Then
          logo_next_avail_id = (cmdarray(cmdcount - 1).ce_valuedword)
          logo_next_avail_id = logo_next_avail_id + 1
          Exit Function
      End If
  Else
      logo_next_avail_id = 8192
  End If
End Function
Private Function getcmdmapnum(ByRef cmd_parm_value As String) As Integer
  logo_sub_keyname = "Software\Microsoft\Internet Explorer\Extensions\CmdMapping"
  logo_sub_valuename = cmd_parm_value
  logo_lreturnvalue = 0
  Try
    rk3 = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
  Catch
    Err.Clear()
    logo_lreturnvalue = 2
  End Try
  If rk3 Is Nothing Then
      logo_lreturnvalue = 2
  Else
      logo_lreturnvalue = 0
  End If
  logo_string_ret_length = 8
 If logo_lreturnvalue = 0 Then
   If Me.win2kxp_flag = "Y" Then
     Try
       ie7_value_returned = rk3.GetValue(logo_sub_valuename, "notfound")
     Catch
       Err.Clear()
       logo_lreturnvalue = 2
     End Try
       If ie7_value_returned.ToString = "notfound" Then
           logo_lreturnvalue = 2
       Else
        If logo_lreturnvalue = 0 Then
         If ie7_value_returned.GetType.ToString = "System.Int32" Or _
            ie7_value_returned.GetType.ToString = "System.String" Then
            logo_string_ret_length = ie7_value_returned.ToString.Length
            If logo_string_ret_length >= 1 Then
              If IsNumeric(ie7_value_returned.ToString) Then
                 logo_lreturnvalue = 0
              Else
                 logo_lreturnvalue = 2
              End If
            Else
              logo_lreturnvalue = 2
            End If
         Else
           logo_lreturnvalue = 2
         End If
        End If
       End If
   Else
       If Me.winverflag = "winnt4" Then
         Try
           ie7_value_returned = rk3.GetValue(logo_sub_valuename, "notfound")
         Catch
           Err.Clear()
           logo_lreturnvalue = 2
         End Try
           If ie7_value_returned.ToString = "notfound" Then
               logo_lreturnvalue = 2
           Else
            If logo_lreturnvalue = 0 Then
             If ie7_value_returned.GetType.ToString = "System.Int32" Or _
               ie7_value_returned.GetType.ToString = "System.String" Then
               logo_string_ret_length = ie7_value_returned.ToString.Length
               If logo_string_ret_length >= 1 Then
                 If IsNumeric(ie7_value_returned.ToString) Then
                   logo_lreturnvalue = 0
                 Else
                   logo_lreturnvalue = 2
                 End If
               Else
                 logo_lreturnvalue = 2
               End If
             Else
               logo_lreturnvalue = 2
             End If
            End If
           End If
       Else
          Try
             ie7_value_returned = rk3.GetValue(logo_sub_valuename, "notfound")
          Catch
             Err.Clear()
             logo_lreturnvalue = 2
          End Try
           If ie7_value_returned.ToString = "notfound" Then
               logo_lreturnvalue = 2
           Else
            If logo_lreturnvalue = 0 Then
             If ie7_value_returned.GetType.ToString = "System.Int32" Or _
               ie7_value_returned.GetType.ToString = "System.String" Then
               logo_string_ret_length = ie7_value_returned.ToString.Length
               If logo_string_ret_length >= 1 Then
                 If IsNumeric(ie7_value_returned.ToString) Then
                   logo_lreturnvalue = 0
                 Else
                   logo_lreturnvalue = 2
                 End If
               Else
                 logo_lreturnvalue = 2
               End If
             Else
               logo_lreturnvalue = 2
             End If
            End If
           End If
       End If
   End If
 End If
 If logo_lreturnvalue = 0 Then
    logo_string_returned = ie7_value_returned.ToString
    logo_long_byte = Mid(logo_string_returned, 1, 4)
    logo_string_returned = (Hex(logo_long_byte)).ToString
    logo_long_returned = CInt(logo_long_byte)
    getcmdmapnum = logo_long_returned
 Else
     getcmdmapnum = 0
 End If
 If Not rk3 Is Nothing Then
     rk3.Close()
 End If
End Function
Private Function setcmdmapnum(ByRef cmd_parm_value_set As String, ByRef cmd_parm_double As Integer) As Integer
  logo_lreturnvalue = 0
  logo_subkey = "Software\Microsoft\Internet Explorer\Extensions\CmdMapping"
  logo_string_key_value = cmd_parm_value_set
  Try
    rk3 = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_subkey, True)
  Catch
    Err.Clear()
    logo_lreturnvalue = 2
  End Try
  If rk3 Is Nothing Then
      logo_lreturnvalue = 2
  Else
      logo_lreturnvalue = 0
  End If
  If logo_lreturnvalue = 0 Then
      logo_subkey = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Extensions\CmdMapping"
      logo_cmdvalue_obj = cmd_parm_double
      If Not rk3 Is Nothing Then
        Try
          rk3.SetValue(cmd_parm_value_set, logo_cmdvalue_obj, RegistryValueKind.DWord)
        Catch
          Err.Clear()
          create_prob_flag = True
        End Try
      End If
  End If
End Function
Private Sub SetUpKeyValue(ByRef hkey_id As Integer, ByRef subkey As String, ByRef passed_key_value As String, ByRef passed_parm_value As String)
  Dim logo_string_parm_value As String
  logo_string_key_value = passed_key_value
  logo_string_parm_value = passed_parm_value & vbNullChar
  If hkey_id = HKEY_LOCAL_MACHINE Then
    Try
      rk2 = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(subkey, True)
    Catch
      Err.Clear()
      create_prob_flag = True
    End Try
      If Not rk2 Is Nothing Then
        Try
          rk2.SetValue(logo_string_key_value, logo_string_parm_value)
        Catch
          Err.Clear()
          create_prob_flag = True
        End Try
      End If
  Else
      rk2 = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(subkey, True)
      If Not rk2 Is Nothing Then
        Try
          rk2.SetValue(logo_string_key_value, logo_string_parm_value)
        Catch
          Err.Clear()
          create_prob_flag = True
        End Try
      End If
  End If
End Sub
Private Sub build_new_buttons()
  If logo_last_button_value < 254 Then
      logo_last_button_value = logo_last_button_value + 1
  Else
      logo_last_button_value = 0
  End If
  logo_last_button_value_2 = &H3S
  ReDim logo_default_toolbar_byte(logo_string_ret_length + 84)
  logo_default_toolbar_map(logo_string_ret_length + 0) = CByte(logo_last_button_value)
  logo_default_toolbar_map(logo_string_ret_length + 1) = CByte(logo_last_button_value_2)
  logo_default_toolbar_map(logo_string_ret_length + 2) = &H0S
  logo_default_toolbar_map(logo_string_ret_length + 3) = &H0S
  logo_default_toolbar_map(logo_string_ret_length + 4) = &H80S
  logo_default_toolbar_map(logo_string_ret_length + 5) = &H69S
  logo_default_toolbar_map(logo_string_ret_length + 6) = &H79S
  logo_default_toolbar_map(logo_string_ret_length + 7) = &H1ES
  logo_default_toolbar_map(logo_string_ret_length + 8) = &HC5S
  logo_default_toolbar_map(logo_string_ret_length + 9) = &H9CS
  logo_default_toolbar_map(logo_string_ret_length + 10) = &HD1S
  logo_default_toolbar_map(logo_string_ret_length + 11) = &H11S
  logo_default_toolbar_map(logo_string_ret_length + 12) = &HA8S
  logo_default_toolbar_map(logo_string_ret_length + 13) = &H3FS
  logo_default_toolbar_map(logo_string_ret_length + 14) = &H0S
  logo_default_toolbar_map(logo_string_ret_length + 15) = &HC0S
  logo_default_toolbar_map(logo_string_ret_length + 16) = &H4FS
  logo_default_toolbar_map(logo_string_ret_length + 17) = &HC9S
  logo_default_toolbar_map(logo_string_ret_length + 18) = &H9DS
  logo_default_toolbar_map(logo_string_ret_length + 19) = &H61S
  logo_default_toolbar_map(logo_string_ret_length + 22) = &H0S
  logo_default_toolbar_map(logo_string_ret_length + 23) = &H0S
  logo_default_toolbar_map(logo_string_ret_length + 24) = &H4S
  logo_default_toolbar_map(logo_string_ret_length + 25) = &H0S
  logo_default_toolbar_map(logo_string_ret_length + 26) = &H0S
  logo_default_toolbar_map(logo_string_ret_length + 27) = &H0S
  If logo_last_button_value < 254 Then
      logo_last_button_value = logo_last_button_value + 1
  Else
      logo_last_button_value = 0
  End If
  dest_string = Hex(logo_guid_map_num_2)
  dest_hi_string = Mid(dest_string, 3, 2)
  dest_lo_string = Mid(dest_string, 1, 2)
  dest_lo_value = hex2lng(dest_hi_string)
  dest_hi_value = hex2lng(dest_lo_string)
  For logo_i = 0 To logo_string_ret_length - 1
      logo_default_toolbar_byte(logo_i) = logo_default_toolbar_map(logo_i)
  Next
  For logo_i = logo_string_ret_length To logo_string_ret_length + 54
      logo_default_toolbar_byte(logo_i) = CByte(logo_default_toolbar_map(logo_i))
  Next
  logo_default_toolbar_byte(logo_string_ret_length + 20) = CByte(dest_lo_value)
  logo_default_toolbar_byte(logo_string_ret_length + 21) = CByte(dest_hi_value)
  logo_string_ret_length = logo_string_ret_length + 27
  logo_subkey = "Software\Microsoft\Internet Explorer\Toolbar"
  logo_string_key_value = "{1E796980-9CC5-11D1-A83F-00C04FC99D61}"
  logo_string_key_value = Mid(logo_string_key_value, 1, 38)
  Try
    rk2 = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_subkey, True)
  Catch
    Err.Clear()
    create_prob_flag = True
  End Try
  If Not rk2 Is Nothing Then
    Try
      rk2.SetValue("{1E796980-9CC5-11D1-A83F-00C04FC99D61}", logo_default_toolbar_byte, RegistryValueKind.Binary)
    Catch
      Err.Clear()
      create_prob_flag = True
    End Try
    Try
      rk2.Close()
    Catch
      Err.Clear()
    End Try
  End If
  logo_button_installed = True
End Sub
Private Sub evaluate_buttons()
  For logo_i = 0 To logo_string_ret_length - 1
      logo_default_toolbar_map(logo_i) = logo_default_toolbar_byte(logo_i)
  Next
  toolbar_cmd_count = 1
  logo_current_button_num = 4
  logo_num_of_separators = 0
  For logo_i = 1 To logo_num_of_toolbar_items
      If logo_default_toolbar_byte(logo_current_button_num) = 255 Or _
         logo_default_toolbar_byte(logo_current_button_num) = 254 Then
          If logo_default_toolbar_byte(logo_current_button_num + 1) = 255 Or _
          logo_default_toolbar_byte(logo_current_button_num + 1) = 254 Then
              logo_default_separator = CInt(((logo_current_button_num - 4) / 28))
              logo_separator_found = True
              logo_num_of_separators = logo_num_of_separators + 1
              If logo_num_of_separators = 1 Then
                  logo_pos_sep_1 = logo_default_separator
              End If
              If logo_num_of_separators = 2 Then
                  logo_pos_sep_2 = logo_default_separator
              End If
          End If
      End If
      If logo_default_toolbar_byte(logo_current_button_num + 1) < 254 Then
          toolbar_cmd_array(toolbar_cmd_count).tm_value_hi_dword = logo_default_toolbar_byte(logo_current_button_num + 20)
          toolbar_cmd_array(toolbar_cmd_count).tm_value_lo_dword = logo_default_toolbar_byte(logo_current_button_num + 21)
          dest_hi_value = logo_default_toolbar_byte(logo_current_button_num + 20)
          dest_hi_string = CStr(logo_default_toolbar_byte(logo_current_button_num + 20))
          dest_lo_string = CStr(logo_default_toolbar_byte(logo_current_button_num + 21))
          If Len(dest_lo_string) < 2 Then
              dest_lo_string = "0" & dest_lo_string
              dest_lo_value = hex2lng2(dest_lo_string)
              dest_hi_value = dest_hi_value + dest_lo_value
              toolbar_string = Hex(dest_hi_value)
          Else
              dest_hi_string = CStr(logo_default_toolbar_byte(logo_current_button_num + 20))
              dest_hi_string = Hex(CInt(dest_hi_string))
              dest_lo_string = CStr(logo_default_toolbar_byte(logo_current_button_num + 21))
              dest_lo_string = Hex(CInt(dest_lo_string))
              toolbar_string = dest_lo_string & dest_hi_string
          End If
          If Len(dest_hi_string) < 2 Then
              dest_hi_string = "0" & dest_hi_string
          End If
          If Len(dest_lo_string) < 2 Then
              dest_lo_string = "0" & dest_lo_string
          End If
          toolbar_string = dest_lo_string & dest_hi_string
          toolbar_cmd_array(toolbar_cmd_count).tm_value_lo_dword = logo_default_toolbar_byte(logo_current_button_num + 21)
          toolbar_cmd_array(toolbar_cmd_count).tm_value_hi_dword = logo_default_toolbar_byte(logo_current_button_num + 20)
          toolbar_cmd_array(toolbar_cmd_count).tm_valuehex = toolbar_string
          toolbar_cmd_count = toolbar_cmd_count + 1
      End If
      logo_current_button_num = logo_current_button_num + 28
  Next
  analyze_buttons()
End Sub
Private Sub analyze_buttons()
  logo_first_mapid_pos = 0
  If logo_button_installed = True Then
      If IE7_flag = False Then
          success_install_msg()
          Exit Sub
      End If
  End If
  If logo_cmdmapping_flag = True Then
      match_map_ids()
  End If
  If logo_use_mapid_logic = True Then
      If logo_first_mapid_pos > 0 Then
          If logo_first_mapid_pos < logo_default_target_pos Then
              logo_target_button_pos = logo_first_mapid_pos
              insert_new_button()
              If IE7_flag = False Then
                  success_install_msg()
              End If
              Exit Sub
          Else
              logo_target_button_pos = logo_default_target_pos
              insert_new_button()
              If IE7_flag = False Then
                  success_install_msg()
              End If
              Exit Sub
          End If
      Else
          If logo_mapid_found_flag = True Then
              If logo_first_mapid_pos < logo_default_target_pos Then
                  logo_target_button_pos = logo_first_mapid_pos
                  insert_new_button()
                  If IE7_flag = False Then
                      success_install_msg()
                  End If
                  Exit Sub
              Else
                  logo_target_button_pos = logo_default_target_pos
                  insert_new_button()
                  If IE7_flag = False Then
                      success_install_msg()
                  End If
                  Exit Sub
              End If
          Else
              logo_target_button_pos = logo_default_target_pos
              insert_new_button()
              If IE7_flag = False Then
                  success_install_msg()
              End If
              Exit Sub
          End If
      End If
  Else
      logo_target_button_pos = logo_default_target_pos
      insert_new_button()
      If IE7_flag = False Then
          success_install_msg()
      End If
      Exit Sub
  End If
End Sub
Private Sub insert_new_button()
  For logo_i = 0 To logo_string_ret_length - 1
      logo_default_toolbar_map(logo_i) = logo_default_toolbar_byte(logo_i)
  Next
  ReDim logo_default_toolbar_byte(logo_string_ret_length + 84)
  If logo_target_button_pos > 1 Then
      If logo_num_of_toolbar_items > logo_default_target_pos Then
          logo_button_difference = logo_num_of_toolbar_items - logo_target_button_pos
          logo_num_bytes_to_move = logo_button_difference * 28
          logo_start_pos = ((logo_target_button_pos * 28) + 4)
      Else
          If logo_num_of_toolbar_items < logo_default_target_pos Then
              logo_target_button_pos = logo_num_of_toolbar_items
              logo_button_difference = logo_num_of_toolbar_items
              logo_num_bytes_to_move = logo_button_difference * 28
              logo_start_pos = (((logo_target_button_pos * 28) + 4) - 1)
          Else
              If logo_num_of_toolbar_items = logo_default_target_pos Then
                  logo_target_button_pos = logo_num_of_toolbar_items
                  logo_button_difference = logo_num_of_toolbar_items
                  logo_num_bytes_to_move = logo_button_difference * 28
                  logo_start_pos = (((logo_target_button_pos * 28) + 4) - 1)
              End If
          End If
      End If
  Else
      If logo_target_button_pos = 1 Then
          logo_button_difference = logo_num_of_toolbar_items - 1
          logo_num_bytes_to_move = logo_button_difference * 28
          logo_start_pos = (logo_target_button_pos * 28) + 4
      End If

      If logo_target_button_pos = 0 Then
          logo_button_difference = logo_num_of_toolbar_items
          logo_num_bytes_to_move = logo_button_difference * 28
          logo_start_pos = (logo_target_button_pos * 28) + 4
      End If
  End If
  For logo_i = 0 To logo_start_pos
      logo_default_toolbar_byte(logo_i) = logo_default_toolbar_map(logo_i)
  Next
  For logo_i = logo_start_pos To (logo_start_pos + logo_num_bytes_to_move)
      logo_default_toolbar_byte(logo_i + 28) = logo_default_toolbar_map(logo_i)
  Next
  logo_start_byte_pos = (((logo_target_button_pos * 28) + 4) - 1)
  logo_last_button_value = logo_default_toolbar_byte(4)
  logo_last_button_value_2 = logo_default_toolbar_byte(5)
  If logo_last_button_value < 254 Then
      logo_seq_button_value = logo_default_toolbar_byte(4)
      logo_seq_button_value_2 = logo_default_toolbar_byte(5)
  Else
      logo_seq_button_value = &HF6S
      logo_seq_button_value_2 = &H3S
  End If
  logo_start_byte_pos = ((logo_target_button_pos * 28) + 4)
  logo_default_toolbar_byte(logo_start_byte_pos + 2) = &H0S
  logo_default_toolbar_byte(logo_start_byte_pos + 3) = &H0S
  logo_default_toolbar_byte(logo_start_byte_pos + 4) = &H80S
  logo_default_toolbar_byte(logo_start_byte_pos + 5) = &H69S
  logo_default_toolbar_byte(logo_start_byte_pos + 6) = &H79S
  logo_default_toolbar_byte(logo_start_byte_pos + 7) = &H1ES
  logo_default_toolbar_byte(logo_start_byte_pos + 8) = &HC5S
  logo_default_toolbar_byte(logo_start_byte_pos + 9) = &H9CS
  logo_default_toolbar_byte(logo_start_byte_pos + 10) = &HD1S
  logo_default_toolbar_byte(logo_start_byte_pos + 11) = &H11S
  logo_default_toolbar_byte(logo_start_byte_pos + 12) = &HA8S
  logo_default_toolbar_byte(logo_start_byte_pos + 13) = &H3FS
  logo_default_toolbar_byte(logo_start_byte_pos + 14) = &H0S
  logo_default_toolbar_byte(logo_start_byte_pos + 15) = &HC0S
  logo_default_toolbar_byte(logo_start_byte_pos + 16) = &H4FS
  logo_default_toolbar_byte(logo_start_byte_pos + 17) = &HC9S
  logo_default_toolbar_byte(logo_start_byte_pos + 18) = &H9DS
  logo_default_toolbar_byte(logo_start_byte_pos + 19) = &H61S
  dest_string = Hex(logo_guid_map_num_2)
  dest_hi_string = Mid(dest_string, 3, 2)
  dest_lo_string = Mid(dest_string, 1, 2)
  dest_lo_value = hex2lng(dest_hi_string)
  dest_hi_value = hex2lng(dest_lo_string)
  logo_default_toolbar_byte(logo_start_byte_pos + 20) = CByte(dest_lo_value)
  logo_default_toolbar_byte(logo_start_byte_pos + 21) = CByte(dest_hi_value)
  logo_default_toolbar_byte(logo_start_byte_pos + 22) = &H0S
  logo_default_toolbar_byte(logo_start_byte_pos + 23) = &H0S
  logo_default_toolbar_byte(logo_start_byte_pos + 24) = &H4S
  logo_default_toolbar_byte(logo_start_byte_pos + 25) = &H0S
  logo_default_toolbar_byte(logo_start_byte_pos + 26) = &H0S
  logo_default_toolbar_byte(logo_start_byte_pos + 27) = &H0S
  For logo_i = 0 To logo_num_of_toolbar_items
   If logo_default_toolbar_byte(logo_string_ret_length + 0) = logo_default_toolbar_byte(logo_string_ret_length + 20) Then
     If logo_default_toolbar_byte(logo_string_ret_length + 1) = logo_default_toolbar_byte(logo_string_ret_length + 21) Then
              'continue'
     End If
   Else
     logo_start_byte_pos = (logo_i * 28) + 4
     logo_last_button_value = logo_default_toolbar_byte(logo_start_byte_pos)
     logo_last_button_value_2 = logo_default_toolbar_byte(logo_start_byte_pos + 1)
     If logo_last_button_value = 255 Or _
        logo_last_button_value = 254 Then
         If logo_last_button_value_2 = 255 Or _
            logo_last_button_value_2 = 254 Then
         End If
     End If
     If logo_last_button_value < 256 Then
         If logo_last_button_value_2 < 254 Then
            If logo_last_button_value_2 = 3 Then
               If logo_seq_button_value_2 = 4 Then
                  logo_last_button_value_2 = 4
                  logo_default_toolbar_byte(logo_start_byte_pos + 1) = CByte(logo_last_button_value_2)
               End If
            End If
            If logo_last_button_value_2 = 4 Then
               If logo_seq_button_value_2 = 3 Then
                  logo_last_button_value_2 = 3
                  logo_default_toolbar_byte(logo_start_byte_pos + 1) = CByte(logo_last_button_value_2)
               End If
            End If
            If logo_last_button_value_2 = 0 Then
               logo_last_button_value_2 = logo_seq_button_value_2
               logo_default_toolbar_byte(logo_start_byte_pos + 1) = CByte(logo_seq_button_value_2)
            End If
            If logo_seq_button_value_2 = 4 Then
                logo_last_button_value_2 = 4
                logo_default_toolbar_byte(logo_start_byte_pos + 1) = CByte(logo_last_button_value_2)
            End If
            logo_default_toolbar_byte(logo_start_byte_pos) = CByte(logo_seq_button_value)
         End If
         logo_seq_button_value = logo_seq_button_value + 1
         If logo_seq_button_value > 254 Then
            logo_seq_button_value = 0
            If logo_seq_button_value_2 = 3 Then
               logo_seq_button_value_2 = 4
            End If
         End If
      End If
    End If
  Next
  logo_string_ret_length = logo_string_ret_length + 28
  logo_subkey = "Software\Microsoft\Internet Explorer\Toolbar"
  logo_string_key_value = "{1E796980-9CC5-11D1-A83F-00C04FC99D61}"
  logo_string_key_value = Mid(logo_string_key_value, 1, 38)
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_subkey, True)
  Catch
    Err.Clear()
    create_prob_flag = True
  End Try
  If rk Is Nothing Then
      logo_lreturnvalue = 2
  Else
      logo_lreturnvalue = 0
  End If
  ReDim Preserve logo_default_toolbar_byte(logo_string_ret_length)
  If logo_lreturnvalue = 0 Then
   Try
     rk.SetValue("{1E796980-9CC5-11D1-A83F-00C04FC99D61}", logo_default_toolbar_byte, RegistryValueKind.Binary)
   Catch
     Err.Clear()
     create_prob_flag = True
   End Try
   logo_button_installed = True
  End If
  If Not rk Is Nothing Then
   rk.Close()
  End If
End Sub
Private Sub check_nextid()
  logo_lreturnvalue = 0
  logo_nextid_missing_flag = False     '01/24/2011'
  logo_current_nextid_inuse = False    '01/24/2011'
  logo_sub_keyname = "Software\Microsoft\Internet Explorer\Extensions\CmdMapping"
  logo_sub_valuename = "NextId"
  mapcount = 0
  Try
    rk2 = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
  Catch
    Err.Clear()
    create_prob_flag = True
  End Try
  If rk2 Is Nothing Then
      logo_lreturnvalue = 2
  Else
      logo_lreturnvalue = 0
  End If
  If logo_lreturnvalue = 0 Then
    Try
      ie7_value_returned = rk2.GetValue(logo_sub_valuename, "notfound")
    Catch
      Err.Clear()
      logo_lreturnvalue = 2
    End Try
      If ie7_value_returned.ToString = "notfound" Then
          logo_lreturnvalue = 2
          logo_nextid_missing_flag = True
      Else
        logo_lreturnvalue = 0
        If ie7_value_returned.GetType.ToString = "System.Int32" Then
         logo_string_ret_length = ie7_value_returned.ToString.Length
         If logo_string_ret_length >= 1 Then
          logo_long_returned = CInt(ie7_value_returned.ToString)
          logo_next_avail_id = logo_long_returned
          If logo_next_avail_id < 8192 Or _
             logo_next_avail_id > 8702 Then
               logo_nextid_missing_flag = True
          Else                                                '01/24/2011'
             logo_current_nextid_value = logo_next_avail_id   '01/24/2011'
          End If
         Else
          logo_nextid_missing_flag = True
         End If
        Else
         logo_nextid_missing_flag = True
        End If
      End If
     If logo_lreturnvalue = 2 _
       Or logo_install_activity = False _
       Or logo_nextid_missing_flag = True Then
      logo_lreturnvalue = 0      '01/24/2011'
      If rk2 Is Nothing Then
         logo_rqve_lpcvalues = 0
      Else
         logo_rqve_lpcvalues = rk2.ValueCount
      End If
      If logo_rqve_lpcvalues = 0 Then
         logo_next_avail_id = 8192
         logo_return_value = setcmdmapnum("NextId", logo_next_avail_id)
         logo_nextid_missing_flag = False
      End If
      If logo_rqve_lpcvalues > 0 Then
          rev_dwindex = 0
          rev_hkey = hKey
          rev_lpvaluename = Space(200)
          rev_lpcbvaluename = 200
          If logo_rqve_lpcvalues > 1 Then
            Try
               rk2.GetValueNames()
            Catch
              Err.Clear()
              create_prob_flag = True
            End Try
          For Each valueName As String In rk2.GetValueNames()
           rev_lpvaluename = Space(200)
           rev_lpcbvaluename = 200
           If valueName.GetType.ToString = "System.String" Then   '01/24/2011'
              rev_lpvaluename = valueName.ToString                '01/24/2011'
           End If                                                 '01/24/2011'
          If Mid(rev_lpvaluename, 1, 1) = "{" Then
             cmdmapping_array(mapcount).cmd_valuename = rev_lpvaluename
            Try
              ie7_value_returned = rk2.GetValue(valueName, "notfound")
            Catch
              Err.Clear()
              logo_lreturnvalue = 2
            End Try
             If ie7_value_returned.ToString = "notfound" Or _
                logo_lreturnvalue = 2 Then
             Else
               If ie7_value_returned.GetType.ToString = "System.Int32" Or _
                  ie7_value_returned.GetType.ToString = "System.String" Then
                logo_string_ret_length = ie7_value_returned.ToString.Length
                If logo_string_ret_length >= 1 Then
                 If IsNumeric(ie7_value_returned.ToString) Then
                   cmdmapping_array(mapcount).cmd_valuedword = _
                   CInt(rk2.GetValue(valueName).ToString)
                   mapcount = mapcount + 1
                 End If
                End If
               End If
             End If
          End If
        Next
     If mapcount > 0 Then
       ReDim cmdmapping_sort_array(mapcount - 1)
       For rev_dwindex = 0 To mapcount - 1
         If cmdmapping_array(rev_dwindex).cmd_valuedword > 0 Then
           cmdmapping_sort_array(rev_dwindex) = _
             cmdmapping_array(rev_dwindex).cmd_valuedword.ToString
         End If
       Next
      Dim mycomparer3 As New myReverserClass()
      Array.Sort(cmdmapping_sort_array, mycomparer3)
      logo_high_dblword_value = CInt(cmdmapping_sort_array(mapcount - 1))
     End If
      rev_lpvaluename = Space(200)
      rev_lpcbvaluename = 200
      rev_dwindex = logo_rqve_lpcvalues - 1
      If logo_high_dblword_value >= 0 Then
          If logo_high_dblword_value < 8702 And logo_high_dblword_value >= 8192 Then
              logo_next_avail_id = logo_high_dblword_value + 1
              logo_return_value = setcmdmapnum("NextId", logo_next_avail_id)
              logo_nextid_missing_flag = False
          Else
              If logo_install_activity = False Then
               ReDim cmdarray(90)
               ReDim cmdarray_sort(90)
               For rev_dwindex = 0 To mapcount - 1
                If cmdmapping_array(rev_dwindex).cmd_valuedword > 0 Then
                 cmdarray(rev_dwindex).ce_valuename = _
                  cmdmapping_array(rev_dwindex).cmd_valuename
                 cmdarray(rev_dwindex).ce_valuedword = _
                  cmdmapping_array(rev_dwindex).cmd_valuedword
                End If
               Next
               For rev_dwindex = 0 To mapcount - 1
                 current_cmdmap_id = CInt(cmdmapping_sort_array(rev_dwindex))
                 For cmdmap_loop = 0 To mapcount - 1
                  If current_cmdmap_id = cmdmapping_array(cmdmap_loop).cmd_valuedword Then
                    cmdarray_sort(rev_dwindex).ce_valuedword = _
                    cmdmapping_array(cmdmap_loop).cmd_valuedword
                    cmdarray_sort(rev_dwindex).ce_valuename = _
                    cmdmapping_array(cmdmap_loop).cmd_valuename
                    Exit For
                  End If
                 Next
               Next
               cmdcount = mapcount
               For cmdmap_loop = 0 To mapcount - 1
                cmdmapping_array(cmdmap_loop).cmd_valuedword = _
                 cmdarray_sort(rev_dwindex).ce_valuedword
                cmdmapping_array(cmdmap_loop).cmd_valuename = _
                 cmdarray_sort(rev_dwindex).ce_valuename
               Next
               find_next_avail_id()
               logo_return_value = setcmdmapnum("NextId", logo_next_avail_id)
               logo_nextid_missing_flag = False
              End If
            End If
          End If
        Else
         If logo_rqve_lpcvalues = 1 Then
            Try
              rk2.GetValueNames()
            Catch
              Err.Clear()
              create_prob_flag = True
            End Try
            For Each valueName As String In rk2.GetValueNames()
             If valueName = "NextId" Then
              rev_lpvaluename = Space(200)
              rev_lpcbvaluename = 200
              If logo_next_avail_id < 8192 Or _
               logo_next_avail_id > 8702 Then
               logo_next_avail_id = 8192
               logo_return_value = setcmdmapnum("NextId", logo_next_avail_id)
               logo_nextid_missing_flag = False
              End If
             Else
               If rk2.GetValue(valueName).GetType.ToString = "System.Int32" Then
                logo_next_avail_id = CInt(rk2.GetValue(valueName).ToString)
                If logo_next_avail_id >= 8192 And logo_next_avail_id <= 8699 Then
                  logo_next_avail_id = logo_next_avail_id + 1
                  logo_cmdmap_spread_above = 8702 - logo_next_avail_id   '01/27/2011'
                  logo_cmdmap_spread_below = logo_next_avail_id - 8192   '01/27/2011'
                  If logo_cmdmap_spread_above >= logo_cmdmap_spread_below Then
                    logo_next_avail_id = logo_cmdmap_spread_above        '01/27/2011'
                    logo_nextid_spread_changed = True                    '01/27/2011'
                  Else
                    logo_next_avail_id = 8192
                    logo_nextid_spread_changed = True
                  End If
                Else
                  logo_next_avail_id = 8192
                End If
                logo_return_value = setcmdmapnum("NextId", logo_next_avail_id)
                logo_nextid_missing_flag = False
               End If
             End If
            Next
         End If
        End If
       End If
     Else
         'continue'
      If rk2 Is Nothing Then
         logo_rqve_lpcvalues = 0
      Else
         logo_rqve_lpcvalues = rk2.ValueCount
      End If
      If logo_rqve_lpcvalues = 0 Then
         logo_next_avail_id = 8192
         logo_return_value = setcmdmapnum("NextId", logo_next_avail_id)
         logo_nextid_missing_flag = False
      End If
      If logo_rqve_lpcvalues > 0 Then
          rev_dwindex = 0
          rev_hkey = hKey
          rev_lpvaluename = Space(200)
          rev_lpcbvaluename = 200
          If logo_rqve_lpcvalues > 1 Then
            Try
               rk2.GetValueNames()
            Catch
              Err.Clear()
              create_prob_flag = True
            End Try
          For Each valueName As String In rk2.GetValueNames()
           rev_lpvaluename = Space(200)
           rev_lpcbvaluename = 200
           If valueName.GetType.ToString = "System.String" Then   '01/24/2011'
              rev_lpvaluename = valueName.ToString                '01/24/2011'
           End If                                                 '01/24/2011'
          If Mid(rev_lpvaluename, 1, 1) = "{" Then
             cmdmapping_array(mapcount).cmd_valuename = rev_lpvaluename
            Try
              ie7_value_returned = rk2.GetValue(valueName, "notfound")
            Catch
              Err.Clear()
              logo_lreturnvalue = 2
            End Try
             If ie7_value_returned.ToString = "notfound" Or _
                logo_lreturnvalue = 2 Then
             Else
               If ie7_value_returned.GetType.ToString = "System.Int32" Or _
                  ie7_value_returned.GetType.ToString = "System.String" Then
                logo_string_ret_length = ie7_value_returned.ToString.Length
                If logo_string_ret_length >= 1 Then
                 If IsNumeric(ie7_value_returned.ToString) Then
                   cmdmapping_array(mapcount).cmd_valuedword = _
                   CInt(rk2.GetValue(valueName).ToString)
                   mapcount = mapcount + 1
                 End If
                End If
               End If
             End If
          End If
        Next
     If mapcount > 0 Then
       ReDim cmdmapping_sort_array(mapcount - 1)
       For rev_dwindex = 0 To mapcount - 1
         If cmdmapping_array(rev_dwindex).cmd_valuedword > 0 Then
           cmdmapping_sort_array(rev_dwindex) = _
             cmdmapping_array(rev_dwindex).cmd_valuedword.ToString
           'check to see if the NextId value is already being used - 01/24/2011'
           If logo_current_nextid_value > 0 And _
        logo_current_nextid_value >= 8192 And _
          logo_current_nextid_value <= 8702 Then
             If cmdmapping_sort_array(rev_dwindex) = logo_current_nextid_value Then
      'NextId value is already being used'
               logo_current_nextid_inuse = True     '01/24/2011'
             End If
           End If
         End If
       Next
      Dim mycomparer3 As New myReverserClass()
      Array.Sort(cmdmapping_sort_array, mycomparer3)
      logo_high_dblword_value = CInt(cmdmapping_sort_array(mapcount - 1))
     End If
      rev_lpvaluename = Space(200)
      rev_lpcbvaluename = 200
      rev_dwindex = logo_rqve_lpcvalues - 1
      If logo_high_dblword_value >= 0 Then
           If logo_high_dblword_value < 8702 And logo_high_dblword_value >= 8192 Then
             If mapcount > 1 Then
               If logo_current_nextid_inuse = True Then         '01/24/2011'
                 logo_next_avail_id = logo_high_dblword_value + 1
                 logo_return_value = setcmdmapnum("NextId", logo_next_avail_id)
                 logo_nextid_missing_flag = False
               End If
             Else
               If mapcount = 1 Then
                  logo_cmdmap_spread_above = 8702 - logo_high_dblword_value
                  logo_cmdmap_spread_below = logo_high_dblword_value - 8192
                  If logo_cmdmap_spread_above >= logo_cmdmap_spread_below Then
                    If logo_high_dblword_value > 8192 Then                             '03/01/2013'                      
                      logo_next_avail_id = logo_high_dblword_value + 1                 '03/01/2013'
                      logo_return_value = setcmdmapnum("NextId", logo_next_avail_id)
                      logo_nextid_missing_flag = False
                      logo_nextid_spread_changed = True                                '01/27/2011'
                     Else
                      logo_next_avail_id = 8193                                        '03/01/2013'
                      logo_return_value = setcmdmapnum("NextId", logo_next_avail_id)   '03/01/2013'
                      logo_nextid_missing_flag = False                                 '03/01/2013'
                      logo_nextid_spread_changed = True                                '03/01/2013'
                    End If                                                             '03/01/2013'
                  Else
                    logo_next_avail_id = 8192
                    logo_return_value = setcmdmapnum("NextId", logo_next_avail_id)
                    logo_nextid_missing_flag = False
                    logo_nextid_spread_changed = True          '01/27/2011'
                  End If
                End If   'end of check for mapcount = 1'
             End If
          Else
            End If
          End If
      End If
    End If
         'end - 01/24/2011'
     End If
   Else
         'continue'
   End If
  If Not rk2 Is Nothing Then
      rk2.Close()
  End If
End Sub
Private Sub build_default_toolbar()
  ReDim logo_default_toolbar_map(143)
  logo_default_toolbar_map(0) = &H7S
  logo_default_toolbar_map(1) = &H0S
  logo_default_toolbar_map(2) = &H0S
  logo_default_toolbar_map(3) = &H0S
  logo_default_toolbar_map(4) = &HF6S
  logo_default_toolbar_map(5) = &H3S
  logo_default_toolbar_map(6) = &H0S
  logo_default_toolbar_map(7) = &H0S
  logo_default_toolbar_map(8) = &H7ES
  logo_default_toolbar_map(9) = &H69S
  logo_default_toolbar_map(10) = &H79S
  logo_default_toolbar_map(11) = &H1ES
  logo_default_toolbar_map(12) = &HC5S
  logo_default_toolbar_map(13) = &H9CS
  logo_default_toolbar_map(14) = &HD1S
  logo_default_toolbar_map(15) = &H11S
  logo_default_toolbar_map(16) = &HA8S
  logo_default_toolbar_map(17) = &H3FS
  logo_default_toolbar_map(18) = &H0S
  logo_default_toolbar_map(19) = &HC0S
  logo_default_toolbar_map(20) = &H4FS
  logo_default_toolbar_map(21) = &HC9S
  logo_default_toolbar_map(22) = &H9DS
  logo_default_toolbar_map(23) = &H61S
  logo_default_toolbar_map(24) = &H20S
  logo_default_toolbar_map(25) = &H1S
  logo_default_toolbar_map(26) = &H0S
  logo_default_toolbar_map(27) = &H0S
  logo_default_toolbar_map(28) = &H0S
  logo_default_toolbar_map(29) = &H0S
  logo_default_toolbar_map(30) = &H0S
  logo_default_toolbar_map(31) = &H0S
  logo_default_toolbar_map(32) = &HF7S
  logo_default_toolbar_map(33) = &H3S
  logo_default_toolbar_map(34) = &H0S
  logo_default_toolbar_map(35) = &H0S
  logo_default_toolbar_map(36) = &H7ES
  logo_default_toolbar_map(37) = &H69S
  logo_default_toolbar_map(38) = &H79S
  logo_default_toolbar_map(39) = &H1ES
  logo_default_toolbar_map(40) = &HC5S
  logo_default_toolbar_map(41) = &H9CS
  logo_default_toolbar_map(42) = &HD1S
  logo_default_toolbar_map(43) = &H11S
  logo_default_toolbar_map(44) = &HA8S
  logo_default_toolbar_map(45) = &H3FS
  logo_default_toolbar_map(46) = &H0S
  logo_default_toolbar_map(47) = &HC0S
  logo_default_toolbar_map(48) = &H4FS
  logo_default_toolbar_map(49) = &HC9S
  logo_default_toolbar_map(50) = &H9DS
  logo_default_toolbar_map(51) = &H61S
  logo_default_toolbar_map(52) = &H21S
  logo_default_toolbar_map(53) = &H1S
  logo_default_toolbar_map(54) = &H0S
  logo_default_toolbar_map(55) = &H0S
  logo_default_toolbar_map(56) = &H0S
  logo_default_toolbar_map(57) = &H0S
  logo_default_toolbar_map(58) = &H0S
  logo_default_toolbar_map(59) = &H0S
  logo_default_toolbar_map(60) = &HF8S
  logo_default_toolbar_map(61) = &H3S
  logo_default_toolbar_map(62) = &H0S
  logo_default_toolbar_map(63) = &H0S
  logo_default_toolbar_map(64) = &H7ES
  logo_default_toolbar_map(65) = &H69S
  logo_default_toolbar_map(66) = &H79S
  logo_default_toolbar_map(67) = &H1ES
  logo_default_toolbar_map(68) = &HC5S
  logo_default_toolbar_map(69) = &H9CS
  logo_default_toolbar_map(70) = &HD1S
  logo_default_toolbar_map(71) = &H11S
  logo_default_toolbar_map(72) = &HA8S
  logo_default_toolbar_map(73) = &H3FS
  logo_default_toolbar_map(74) = &H0S
  logo_default_toolbar_map(75) = &HC0S
  logo_default_toolbar_map(76) = &H4FS
  logo_default_toolbar_map(77) = &HC9S
  logo_default_toolbar_map(78) = &H9DS
  logo_default_toolbar_map(79) = &H61S
  logo_default_toolbar_map(80) = &H24S
  logo_default_toolbar_map(81) = &H1S
  logo_default_toolbar_map(82) = &H0S
  logo_default_toolbar_map(83) = &H0S
  logo_default_toolbar_map(84) = &H4S
  logo_default_toolbar_map(85) = &H0S
  logo_default_toolbar_map(86) = &H0S
  logo_default_toolbar_map(87) = &H0S
  logo_default_toolbar_map(88) = &HF9S
  logo_default_toolbar_map(89) = &H3S
  logo_default_toolbar_map(90) = &H0S
  logo_default_toolbar_map(91) = &H0S
  logo_default_toolbar_map(92) = &H7ES
  logo_default_toolbar_map(93) = &H69S
  logo_default_toolbar_map(94) = &H79S
  logo_default_toolbar_map(95) = &H1ES
  logo_default_toolbar_map(96) = &HC5S
  logo_default_toolbar_map(97) = &H9CS
  logo_default_toolbar_map(98) = &HD1S
  logo_default_toolbar_map(99) = &H11S
  logo_default_toolbar_map(100) = &HA8S
  logo_default_toolbar_map(101) = &H3FS
  logo_default_toolbar_map(102) = &H0S
  logo_default_toolbar_map(103) = &HC0S
  logo_default_toolbar_map(104) = &H4FS
  logo_default_toolbar_map(105) = &HC9S
  logo_default_toolbar_map(106) = &H9DS
  logo_default_toolbar_map(107) = &H61S
  logo_default_toolbar_map(108) = &H25S
  logo_default_toolbar_map(109) = &H1S
  logo_default_toolbar_map(110) = &H0S
  logo_default_toolbar_map(111) = &H0S
  logo_default_toolbar_map(112) = &H4S
  logo_default_toolbar_map(113) = &H0S
  logo_default_toolbar_map(114) = &H0S
  logo_default_toolbar_map(115) = &H0S
  logo_default_toolbar_map(116) = &HFAS
  logo_default_toolbar_map(117) = &H3S
  logo_default_toolbar_map(118) = &H0S
  logo_default_toolbar_map(119) = &H0S
  logo_default_toolbar_map(120) = &H7ES
  logo_default_toolbar_map(121) = &H69S
  logo_default_toolbar_map(122) = &H79S
  logo_default_toolbar_map(123) = &H1ES
  logo_default_toolbar_map(124) = &HC5S
  logo_default_toolbar_map(125) = &H9CS
  logo_default_toolbar_map(126) = &HD1S
  logo_default_toolbar_map(127) = &H11S
  logo_default_toolbar_map(128) = &HA8S
  logo_default_toolbar_map(129) = &H3FS
  logo_default_toolbar_map(130) = &H0S
  logo_default_toolbar_map(131) = &HC0S
  logo_default_toolbar_map(132) = &H4FS
  logo_default_toolbar_map(133) = &HC9S
  logo_default_toolbar_map(134) = &H9DS
  logo_default_toolbar_map(135) = &H61S
  logo_default_toolbar_map(136) = &H22S
  logo_default_toolbar_map(137) = &H1S
  logo_default_toolbar_map(138) = &H0S
  logo_default_toolbar_map(139) = &H0S
  logo_default_toolbar_map(140) = &H4S
  logo_default_toolbar_map(141) = &H0S
  logo_default_toolbar_map(142) = &H0S
  logo_default_toolbar_map(143) = &H0S
  logo_last_button_value = &HFAS
  logo_last_button_value_2 = &H3S
  logo_default_toolbar_length = 143
  logo_subkey = "Software\Microsoft\Internet Explorer\Toolbar"
  logo_string_key_value = "{1E796980-9CC5-11D1-A83F-00C04FC99D61}"
  logo_string_key_value = Mid(logo_string_key_value, 1, 38)
  ReDim logo_default_toolbar_byte(284)
  For logo_i = 0 To logo_default_toolbar_length
      logo_default_toolbar_byte(logo_i) = CByte(logo_default_toolbar_map(logo_i))
  Next
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_subkey, True)
  Catch
    Err.Clear()
    create_prob_flag = True
  End Try
  If Not rk Is Nothing Then
   Try
    rk.SetValue("{1E796980-9CC5-11D1-A83F-00C04FC99D61}", logo_default_toolbar_map, RegistryValueKind.Binary)
   Catch
     Err.Clear()
     create_prob_flag = True
   End Try
    If logo_default_target_pos > 5 Then
      logo_default_target_pos = 5
    End If
    logo_default_toolbar_built = True
  End If
End Sub
Private Sub create_favorites()
  logo_filename_one_ = (logo_favorite_dir & logo_company_name & ".url")
  If logo_install_favorite = True Then
    Try
      logo_file_number = FreeFile()
      FileOpen(logo_file_number, logo_filename_one_, OpenMode.Output)
      PrintLine(logo_file_number, "[InternetShortcut]")
      PrintLine(logo_file_number, "URL=http://" & logo_company_url)
      PrintLine(logo_file_number, "IconIndex=0")
      PrintLine(logo_file_number, "IconFile=" & logo_app_path & logo_install_module)
      FileClose(logo_file_number)
    Catch
      Err.Clear()
      logo_return_code = 0
    End Try
  End If
End Sub
Private Sub locate_favorites()
  logo_filename_counter = 0
  logo_file_character = ""
  logo_string_ret_length = 200
  logo_lreturnvalue = 0
  logo_fav_retvalue = 0
  logo_sub_keyname = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
  logo_sub_valuename = "Favorites"
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
  Catch
    Err.Clear()
    logo_fav_retvalue = 2
  End Try
  If rk Is Nothing Then
      logo_fav_retvalue = 2
  Else
      logo_fav_retvalue = 0
  End If
  If logo_fav_retvalue = 0 Then
      ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
      If ie7_value_returned.ToString = "notfound" Then
          logo_fav_retvalue = 2
      Else
        If ie7_value_returned.GetType.ToString = "System.String" Then
             logo_string_ret_length = ie7_value_returned.ToString.Length
             If logo_string_ret_length >= 1 Then
               logo_favorite_dir = ie7_value_returned.ToString
               logo_fav_retvalue = 0
             Else
               logo_fav_retvalue = 2
             End If
        Else
          logo_fav_retvalue = 2
        End If
      End If
      If Not rk Is Nothing Then
          rk.Close()
      End If
  End If
  If logo_fav_retvalue <> 0 Then
      If logo_fav_retvalue = ERROR_FILE_NOT_FOUND Then
          logo_fav_retvalue = 0
          logo_sub_keyname = "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
          logo_sub_valuename = "Favorites"
          Try
            rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
          Catch
            Err.Clear()
            logo_fav_retvalue = 2
          End Try
          If rk Is Nothing Then
              logo_fav_retvalue = 2
          Else
              logo_fav_retvalue = 0
          End If
          If logo_fav_retvalue = 0 Then
              ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
              If ie7_value_returned.ToString = "notfound" Then
                  logo_fav_retvalue = 2
              Else
               If ie7_value_returned.GetType.ToString = "System.String" Then
                 logo_string_ret_length = ie7_value_returned.ToString.Length
                 If logo_string_ret_length >= 1 Then
                  logo_favorite_dir = ie7_value_returned.ToString
                  logo_fav_retvalue = 0
                 Else
                  logo_fav_retvalue = 2
                 End If
               Else
                 logo_fav_retvalue = 2
               End If
              End If
              If Not rk Is Nothing Then
                  rk.Close()
              End If
          End If
      End If
  End If
  If logo_fav_retvalue <> 0 Then
      If logo_fav_retvalue = ERROR_FILE_NOT_FOUND Then
          logo_fav_retvalue = 0
          logo_sub_keyname = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
          logo_sub_valuename = "Common Favorites"
          Try
            rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_sub_keyname, False)
          Catch
            Err.Clear()
            create_prob_flag = True
            logo_fav_retvalue = 2
          End Try
          If rk Is Nothing Then
              logo_fav_retvalue = 2
          Else
              logo_fav_retvalue = 0
          End If
          If logo_fav_retvalue = 0 Then
              ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
              If ie7_value_returned.ToString = "notfound" Then
                  logo_fav_retvalue = 2
              Else
                If ie7_value_returned.GetType.ToString = "System.String" Then
                  logo_string_ret_length = ie7_value_returned.ToString.Length
                  If logo_string_ret_length >= 1 Then
                    logo_favorite_dir = ie7_value_returned.ToString
                    logo_fav_retvalue = 0
                  Else
                    logo_fav_retvalue = 2
                  End If
                Else
                  logo_fav_retvalue = 2
                End If
              End If
              If Not rk Is Nothing Then
                  rk.Close()
              End If
          End If
      End If
  End If
  If logo_fav_retvalue <> 0 Then
      If logo_fav_retvalue = ERROR_FILE_NOT_FOUND Then
          logo_fav_retvalue = 0
          logo_sub_keyname = "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
          logo_sub_valuename = "Common Favorites"
          Try
            rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_sub_keyname, False)
          Catch
            Err.Clear()
            create_prob_flag = True
          End Try
          If rk Is Nothing Then
              logo_fav_retvalue = 2
          Else
              logo_fav_retvalue = 0
          End If
          If logo_fav_retvalue = 0 Then
              ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")

              If ie7_value_returned.ToString = "notfound" Then
                  logo_fav_retvalue = 2
              Else
                If ie7_value_returned.GetType.ToString = "System.String" Then
                  logo_string_ret_length = ie7_value_returned.ToString.Length
                  If logo_string_ret_length >= 1 Then
                    logo_favorite_dir = ie7_value_returned.ToString
                    logo_fav_retvalue = 0
                  Else
                    logo_fav_retvalue = 2
                  End If
                Else
                  logo_fav_retvalue = 2
                End If
              End If
              If Not rk Is Nothing Then
                  rk.Close()
              End If
          End If
      End If
  End If
  logo_favorite_path_len = logo_favorite_dir.ToString.Length
  If logo_favorite_path_len >= 1 Then
    If Mid(logo_favorite_dir, logo_favorite_path_len, 1) = "\" Then
      logo_favorite_path = logo_favorite_dir
    Else
      logo_favorite_path = logo_favorite_dir & "\"
    End If
  End If
  logo_favorite_dir = logo_favorite_path
  logo_sub_keyname = "Software\Microsoft\Windows\CurrentVersion"
  logo_sub_valuename = "SystemRoot"
  Try
    rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_sub_keyname, False)
  Catch
    Err.Clear()
    logo_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_lreturnvalue = 2
  Else
      logo_lreturnvalue = 0
  End If
  If logo_lreturnvalue = 0 Then
    Try
      ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
    Catch
      Err.Clear()
      logo_lreturnvalue = 2
    End Try
      logo_systemroot_not_found = False
      If ie7_value_returned.ToString = "notfound" Then
          logo_lreturnvalue = 2
      Else
        If ie7_value_returned.GetType.ToString = "System.String" Then
             logo_string_ret_length = ie7_value_returned.ToString.Length
             If logo_string_ret_length >= 1 Then
               logo_string_filename = ie7_value_returned.ToString
               logo_systemroot_not_found = False
               logo_lreturnvalue = 0
             Else
               logo_lreturnvalue = 2
             End If
        Else
          logo_lreturnvalue = 2
        End If
      End If
      If logo_lreturnvalue <> 0 Then
          logo_systemroot_not_found = True
      End If
  End If
  If logo_systemroot_not_found = True Then
      logo_sub_keyname = "Software\Microsoft\Windows NT\CurrentVersion"
      logo_sub_valuename = "SystemRoot"
      Try
        rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_sub_keyname, False)
      Catch
        Err.Clear()
        logo_lreturnvalue = 2
      End Try
      If rk Is Nothing Then
          logo_lreturnvalue = 2
      Else
          logo_lreturnvalue = 0
      End If
      If logo_lreturnvalue = 0 Then
        Try
          ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
        Catch
          Err.Clear()
          logo_lreturnvalue = 2
        End Try
          logo_systemroot_not_found = False
          If ie7_value_returned.ToString = "notfound" Then
              logo_lreturnvalue = 2
          Else
            If ie7_value_returned.GetType.ToString = "System.String" Then
              logo_string_ret_length = ie7_value_returned.ToString.Length
              If logo_string_ret_length >= 1 Then
                logo_string_filename = ie7_value_returned.ToString
                logo_systemroot_not_found = False
                logo_lreturnvalue = 0
              Else
                logo_lreturnvalue = 2
              End If
            Else
              logo_lreturnvalue = 2
            End If
          End If
          If logo_lreturnvalue <> 0 Then
              logo_systemroot_not_found = True
          End If
      End If
      If Not rk Is Nothing Then
          rk.Close()
      End If
  End If
  logo_system_dir = logo_string_filename & "\system"
End Sub
Private Sub get_programs_location()
  Dim lretval As Integer
  filename_counter = 0
  logo_string_returned = ""
  logo_string_ret_length = 200
  logo_string_filename = ""
  lretval = 0
  source_length = Len(My.Application.Info.DirectoryPath)
  If source_length > 2 Then
    If VB.Left(My.Application.Info.DirectoryPath, 1) = "\" Then
        logo_source_dir = Mid(My.Application.Info.DirectoryPath, 1, source_length - 1)
    Else
        logo_source_dir = My.Application.Info.DirectoryPath
    End If
  End If
  filename_counter = 0
  filechar = ""
  logo_string_returned = ""
  logo_string_ret_length = 200
  skeyname = "Software\Microsoft\Windows\CurrentVersion"
  svaluename = "ProgramFilesDir"
  programs_not_found = False
  Try
    rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(skeyname, False)
  Catch
    Err.Clear()
  End Try
  If rk Is Nothing Then
      lretval = 2
  Else
      lretval = 0
  End If
  If lretval = 0 Then
    Try
      ie7_value_returned = rk.GetValue(svaluename, "notfound")
    Catch
      Err.Clear()
      lretval = 2
    End Try
  End If
  programs_not_found = False
  If ie7_value_returned.ToString = "notfound" Then
      lretval = 2
  Else
      If ie7_value_returned.GetType.ToString = "System.String" Then
         logo_string_ret_length = ie7_value_returned.ToString.Length
         If logo_string_ret_length >= 1 Then
            logo_programs_dir = ie7_value_returned.ToString
            logo_programs_path_len = logo_string_ret_length
            lretval = 0
         End If
      Else
         lretval = 2
      End If
  End If
  If lretval <> 0 Then
      programs_not_found = True
  End If
  If programs_not_found = True Then
      skeyname = "Software\Microsoft\Windows\CurrentVersion"
      svaluename = "ProgramFilesDir"
      Try
        rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(skeyname, False)
      Catch
        Err.Clear()
        lretval = 2
      End Try
      If rk Is Nothing Then
          lretval = 2
      Else
          lretval = 0
      End If
      If lretval = 0 Then
        Try
          ie7_value_returned = rk.GetValue(svaluename, "notfound")
        Catch
          Err.Clear()
          lretval = 2
        End Try
          programs_not_found = False
          If ie7_value_returned.ToString = "notfound" Then
              lretval = 2
          Else
            If ie7_value_returned.GetType.ToString = "System.String" Then
              logo_string_ret_length = ie7_value_returned.ToString.Length
              If logo_string_ret_length >= 1 Then
                logo_programs_dir = ie7_value_returned.ToString
                logo_programs_path_len = logo_string_ret_length
                lretval = 0
              Else
                lretval = 2
              End If
            Else
              lretval = 2
            End If
          End If
          If lretval <> 0 Then
              programs_not_found = True
          End If
      End If
  End If
  If programs_not_found = True Then
      skeyname = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
      svaluename = "Programs"
      Try
        rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(skeyname, False)
      Catch
        Err.Clear()
        lretval = 2
      End Try
      If rk Is Nothing Then
          lretval = 2
      Else
          lretval = 0
      End If
      If lretval = 0 Then
        Try
          ie7_value_returned = rk.GetValue(svaluename, "notfound")
        Catch
          Err.Clear()
          lretval = 2
        End Try
          programs_not_found = False
          If ie7_value_returned.ToString = "notfound" Then
              lretval = 2
          Else
            If ie7_value_returned.GetType.ToString = "System.String" Then
              logo_string_ret_length = ie7_value_returned.ToString.Length
              If logo_string_ret_length >= 1 Then
                logo_programs_dir = ie7_value_returned.ToString
                logo_programs_path_len = logo_string_ret_length
                lretval = 0
              Else
                lretval = 2
              End If
            Else
              lretval = 2
            End If
          End If
          If lretval <> 0 Then
              programs_not_found = True
          End If
      End If
  End If
  If programs_not_found = True Then
      skeyname = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
      svaluename = "ProgramFilesDir"
      Try
        rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(skeyname, False)
      Catch
        Err.Clear()
        lretval = 2
      End Try
      If rk Is Nothing Then
          lretval = 2
      Else
          lretval = 0
      End If
      If lretval = 0 Then
          Try
            ie7_value_returned = rk.GetValue(svaluename, "notfound")
          Catch
            Err.Clear()
            lretval = 2
          End Try
          programs_not_found = False
          If ie7_value_returned.ToString = "notfound" Then
              lretval = 2
          Else
            If ie7_value_returned.GetType.ToString = "System.String" Then
              logo_string_ret_length = ie7_value_returned.ToString.Length
              If logo_string_ret_length >= 1 Then
                logo_programs_dir = ie7_value_returned.ToString
                logo_programs_path_len = logo_string_ret_length
                lretval = 0
              Else
                lretval = 2
              End If
            Else
              lretval = 2
            End If
          End If
          If lretval <> 0 Then
              programs_not_found = True
          End If
      End If
  End If
  If programs_not_found = True Then
      skeyname = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
      svaluename = "Common Programs"
      Try
        rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(skeyname, False)
      Catch
        Err.Clear()
        create_prob_flag = True
        lretval = 2
      End Try
      If rk Is Nothing Then
          lretval = 2
      Else
          lretval = 0
      End If
      If lretval = 0 Then
        Try
          ie7_value_returned = rk.GetValue(svaluename).ToString
        Catch
          Err.Clear()
          lretval = 2
        End Try
          programs_not_found = False
          If ie7_value_returned.ToString = "notfound" Then
              lretval = 2
          Else
            If ie7_value_returned.GetType.ToString = "System.String" Then
              logo_string_ret_length = ie7_value_returned.ToString.Length
              If logo_string_ret_length >= 1 Then
                logo_programs_dir = ie7_value_returned.ToString
                logo_programs_path_len = logo_string_ret_length
                lretval = 0
              Else
                lretval = 2
              End If
            Else
              lretval = 2
            End If
          End If
          If lretval <> 0 Then
              programs_not_found = True
          End If
      End If
  End If
  If Not rk Is Nothing Then
      rk.Close()
  End If
  If programs_not_found = False Then
  'continue'
  End If
End Sub

Private Sub getversion_somb()
  Dim osinfo As OSVERSIONINFO
  Dim returnvalue As Integer
  Me.win_vista_flag = "N"            '11/21/2013'
  Me.win7_flag = "N"                 '11/21/2013'
  Me.win8_flag = "N"      '04/21/2012'
  osinfo.dwOSVersionInfoSize = 148
  osinfo.szCSDVersion = Space(128).ToCharArray
  Try
    returnvalue = GetVersionEx(osinfo)
  Catch
    Err.Clear()
    Me.winverflag = "problem"
  End Try
  If Not Me.winverflag = "problem" Then
    With osinfo
      Select Case .dwPlatformId
          Case 1
              Select Case .dwMinorVersion
                  Case 0                               'Windows 95'                      
                      Me.win2kxp_flag = "N"
                      Me.win2000_flag = "N"
                      Me.winverflag = "win95"
                  Case 10                              'Windows 98'                      
                      Me.win2kxp_flag = "N"
                      Me.win2000_flag = "N"
                      Me.winverflag = "win98"
                  Case 90                              'Windows Me'                      
                      Me.win2kxp_flag = "N"
                      Me.win2000_flag = "N"
                      Me.winverflag = "winme"
              End Select
          Case 2
              Select Case .dwMajorVersion
                  Case 3                               'Windows 3.51'                      
                      Me.win2kxp_flag = "N"
                      Me.win2000_flag = "N"
                      Me.winverflag = "winnt351"
                  Case 4                               'Windows NT 4'                      
                      Me.win2kxp_flag = "N"
                      Me.win2000_flag = "N"
                      Me.winverflag = "winnt4"
                  Case 5                 'either win xp, win xp sp2, or win 2003 server'
                      Select Case .dwMinorVersion
                          Case 0                       'Windows 2000'                              
                              Me.win2kxp_flag = "Y"
                              Me.win2000_flag = "Y"
                              Me.winverflag = "win2000"
                          Case 1                       'Windows XP - this is win xp sp1'                              
                              Me.win2kxp_flag = "Y"
                              Me.win2000_flag = "N"
                              Me.winverflag = "winxp1"                                                       
                          Case 2                       'Win XP SP2 or Windows Server 2003'                              
                              Me.win2kxp_flag = "Y"
                              Me.win2000_flag = "N"
                              Me.winverflag = "win2003"
                      End Select    'end of winxp/win2003 logic'                      
                  Case 6              
                      Select Case .dwMinorVersion     '04/21/2012'
                       ' dwMinorVersion 0 is Windows Vista or Windows Server 2008'
                          Case 0
                            Me.win_vista_flag = "Y"    '11/21/2013'
                       ' dwMinorVersion 1 is Windows 7 or Windows Server 2008 R2'
                          Case 1
                            Me.win7_flag = "Y"         '11/21/2013'
                          Case 2                       'Windows 8 or Windows Server 2012'
                            Me.win8_flag = "Y"                                               '04/21/2012'
                          Case 3                       'Windows 8.1 or Windows Server 2012 R2'
                            Me.win8_flag = "Y"                                               '11/10/2013'                         
                        End Select                     '04/21/2012'
                      Me.win2kxp_flag = "Y"
                      Me.win2000_flag = "N"
                      Me.winverflag = "longhorn"
                      logo_check_UAC_flag = True       '11/10/2013'
                  Case 10
                      Me.win8_flag = "Y"               '04/06/2015'
                      Me.win2kxp_flag = "Y"            '04/06/2015'
                      Me.win2000_flag = "N"            '04/06/2015'
                      Me.winverflag = "longhorn"       '04/06/2015'
                      logo_check_UAC_flag = True       '04/06/2015'
              End Select
      End Select
  End With
 Else
   create_prob_flag = True
 End If
End Sub

Private Sub get_IE_version()
  logo_IE_final_msg = ""                                  '02/17/2013'
  logo_lreturnvalue = 0
  logo_sub_keyname = "Software\Microsoft\Internet Explorer"
'  If Me.winverflag = "longhorn" Then
'if this is vista, 7, 8, or 8.1,'
' check to see if svcVersion exists'
'Vista could be running IE7 still, in which case '
' then [Version] has to be read to retrieve the '
'  value.'
' For all other systems, use the Version Vector lookup that is currently used'

'  If Me.win8_flag = "Y" Then                              '04/21/2012'
'    logo_sub_valuename = "svcVersion"                     '04/21/2012'
'  Else                                                    '04/21/2012'
'    logo_sub_valuename = "Version"                        '04/21/2012'
'  End If                                                  '04/21/2012'
'  Try
'    rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_sub_keyname, False)
'  Catch
'    Err.Clear()
'    logo_lreturnvalue = 2
'  End Try
'  If rk Is Nothing Then
'    logo_lreturnvalue = 2
'  Else
'    logo_lreturnvalue = 0
'  End If
'  logo_string_returned = ""
'  logo_string_ret_length = 200
'  If logo_lreturnvalue = 0 Then
'    ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
'    If ie7_value_returned.ToString = "notfound" Then
'       logo_lreturnvalue = 2
'    Else
'       If ie7_value_returned.GetType.ToString = "System.String" Then
'        logo_string_returned = ie7_value_returned.ToString
'        logo_lreturnvalue = 0
'       End If
'    End If
'  End If

If Me.win_vista_flag = "Y" Or Me.win7_flag = "Y" Then                  '11/26/2013'
 logo_sub_valuename = "svcVersion"
 Try
    rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_sub_keyname, False)
  Catch
    Err.Clear()
    logo_lreturnvalue = 2
 End Try
 If rk Is Nothing Then
    logo_lreturnvalue = 2
 Else
    logo_lreturnvalue = 0
 End If
  logo_string_returned = ""
  logo_string_ret_length = 200
  If logo_lreturnvalue = 0 Then
    ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
    If ie7_value_returned.ToString = "notfound" Then
       logo_lreturnvalue = 2
    Else
       If ie7_value_returned.GetType.ToString = "System.String" Then
        logo_string_returned = ie7_value_returned.ToString
        logo_lreturnvalue = 0
       End If
    End If
  End If
  If logo_lreturnvalue <> 0 Then
  'the svcVersion value did not exist - this must be version 7 or 8'
    logo_lreturnvalue = 0
    logo_sub_valuename = "Version"
    logo_string_returned = ""
    logo_string_ret_length = 200
    If logo_lreturnvalue = 0 Then
      ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
      If ie7_value_returned.ToString = "notfound" Then
       logo_lreturnvalue = 2
      Else
       If ie7_value_returned.GetType.ToString = "System.String" Then
        logo_string_returned = ie7_value_returned.ToString
        logo_lreturnvalue = 0
        GoTo determine_IE_version
       End If
      End If
    End If 'end of check for return value 0'
  Else
    GoTo determine_IE_version
  End If
' alternative check through Version Vector'
 GoTo check_version_vector
End If 'end of check for win_vista_flag = true'

If Me.win8_flag = "Y" Then
  logo_sub_valuename = "svcVersion"
Else
  logo_sub_valuename = "Version"
End If
Try
    rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_sub_keyname, False)
  Catch
    Err.Clear()
    logo_lreturnvalue = 2
End Try
If rk Is Nothing Then
    logo_lreturnvalue = 2
Else
    logo_lreturnvalue = 0
End If
  logo_string_returned = ""
  logo_string_ret_length = 200
  If logo_lreturnvalue = 0 Then
    ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
    If ie7_value_returned.ToString = "notfound" Then
       logo_lreturnvalue = 2
    Else
       If ie7_value_returned.GetType.ToString = "System.String" Then
        logo_string_returned = ie7_value_returned.ToString
        logo_lreturnvalue = 0
      GoTo determine_IE_version
       End If
    End If
End If

check_version_vector:
  If logo_lreturnvalue <> 0 Then
      If Not rk Is Nothing Then
          rk.Close()
      End If
      logo_sub_keyname = "Software\Microsoft\Internet Explorer\Version Vector"
      logo_sub_valuename = "IE"
      Try
        rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_sub_keyname, False)
      Catch
        Err.Clear()
        logo_lreturnvalue = 2
      End Try
      logo_string_returned = ""
      logo_string_ret_length = 200
      If rk Is Nothing Then
         logo_lreturnvalue = 2
      Else
         logo_lreturnvalue = 0
      End If
      If logo_lreturnvalue = 0 Then
       ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
       If ie7_value_returned.ToString = "notfound" Then
           logo_lreturnvalue = 2
       Else
           If ie7_value_returned.GetType.ToString = "System.String" Then
             logo_string_returned = ie7_value_returned.ToString
             logo_lreturnvalue = 0
           End If
       End If
       If Not rk Is Nothing Then
          rk.Close()
       End If
      End If
  End If

determine_IE_version:
  If Len(logo_string_returned.ToString) > 1 Then
    IE7_version_string = Mid(logo_string_returned, 1, 1)
  If IE7_version_string = "6" Then
      IE7_flag = False
      logo_IE_final_msg = "To display the new Toolbar in IE6, select New -> Window from the File Menu, or select Internet Explorer from the Windows Start Menu."   '02/17/2013'
  End If
  If IE7_version_string = "7" Then
      IE7_flag = True
      logo_IE_final_msg = "To display the new Toolbar in IE7, select New Window from the File Menu, or select Internet Explorer from the Windows Start Menu."   '02/17/2013'
  End If
  If IE7_version_string = "8" Then
      IE7_flag = True
      IE8_flag = True
      logo_IE_final_msg = "To display the new Toolbar in IE8, select New Session from the File Menu, or select Internet Explorer from the Windows Start Menu."   '02/17/2013'
  End If
  If IE7_version_string = "9" Then
      IE7_flag = True
      IE8_flag = True
      IE9_flag = True
      logo_IE_final_msg = "To display the new Toolbar in IE9, select New Session from the File Menu, or select Internet Explorer from the Windows Start Menu."   '02/17/2013'
  End If
  If IE7_version_string = "1" Then            '11/17/2013'
    If Mid(logo_string_returned, 1, 2) = "10" Then
      IE7_flag = True                         '11/25/2013'
      IE8_flag = True                         '11/25/2013'
      IE9_flag = True                         '11/25/2013'
      IE10_flag = True                        '11/25/2013'
      logo_IE_final_msg = "To display the new Toolbar in IE10 (desktop mode), select New Session from the File Menu, or select Internet Explorer from the Windows Start Menu."   '02/17/2013'
    End If                                    '11/17/2013'
    If Mid(logo_string_returned, 1, 2) = "11" Then
      IE7_flag = True                         '11/25/2013'
      IE8_flag = True                         '11/25/2013'
      IE9_flag = True                         '11/25/2013'
      IE10_flag = True                        '11/25/2013'
      IE11_flag = True                        '11/26/2013'
      logo_IE_final_msg = "To display the new Toolbar in IE11 (desktop mode), select New Session from the File Menu, or select Internet Explorer from the Windows Start Menu."
    End If                                    '11/17/2013'
  End If                                      '11/17/2013'
  'If IE7_version_string = "10" Or _
  '   IE7_version_string = "1" Then   '04/21/2012'
  '    IE7_flag = True
  '    IE8_flag = True
  '    IE9_flag = True
  '    IE10_flag = True
  '    logo_IE_final_msg = "To display the new Toolbar in IE10 (desktop mode), select New Session from the File Menu, or select Internet Explorer from the Windows Start Menu."   '02/17/2013'
  'End If
 End If
End Sub
'-----------------------------------------------------------------------------------------------------'
Private Sub check_for_prev_install()
  dest_string = Hex(logo_guid_map_num_2)
  dest_hi_string = Mid(dest_string, 3, 2)
  dest_lo_string = Mid(dest_string, 1, 2)
  dest_lo_value = hex2lng(dest_hi_string)
  dest_hi_value = hex2lng(dest_lo_string)
  If logo_string_ret_length > 30 Then
      logo_begin_button_num = logo_default_toolbar_byte(4)
      logo_num_of_toolbar_items = CInt(((logo_string_ret_length - 4) / 28))
    For logo_d = 0 To logo_num_of_toolbar_items - 1
      logo_last_button_seq = (logo_d * 28) + 4
      logo_last_button_value = logo_default_toolbar_byte(logo_last_button_seq + 20)
      logo_last_button_value_2 = logo_default_toolbar_byte(logo_last_button_seq + 21)
      If logo_last_button_value = dest_lo_value Then
          If logo_last_button_value_2 = dest_hi_value Then
              logo_toolbar_prev_install = True
          End If
      End If
    Next
  End If
End Sub
Private Sub Delete_Logo_Key(ByRef logo_dlet_keyname As String, ByRef lPredefinedKey As Integer)
  If lPredefinedKey = HKEY_LOCAL_MACHINE Then
    Try
      rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_dlet_keyname, True)
    Catch
      Err.Clear()
      create_prob_flag = True
    End Try
    If Not rk Is Nothing Then
        If rk.SubKeyCount > 0 Then
            netnewpath = logo_dlet_keyname
            Try
              Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(netnewpath)
            Catch
              Err.Clear()
              create_prob_flag = True
            End Try

            If Not rk Is Nothing Then
              rk.Close()
            End If
        Else
            netnewpath = logo_dlet_keyname
            Try
              Microsoft.Win32.Registry.LocalMachine.DeleteSubKey(netnewpath)
            Catch
              Err.Clear()
              create_prob_flag = True
            End Try

            If Not rk Is Nothing Then
              rk.Close()
            End If
        End If
    End If
  Else
    Try
      rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_dlet_keyname, True)
    Catch
      Err.Clear()
    End Try
    If Not rk Is Nothing Then
      If rk.SubKeyCount > 0 Then
          netnewpath = logo_dlet_keyname
          Try
            Microsoft.Win32.Registry.CurrentUser.DeleteSubKeyTree(netnewpath)
          Catch
            Err.Clear()
            create_prob_flag = True
          End Try

          If Not rk Is Nothing Then
            rk.Close()
          End If
      Else
          netnewpath = logo_dlet_keyname
          Try
            Microsoft.Win32.Registry.CurrentUser.DeleteSubKey(netnewpath)
          Catch
            Err.Clear()
            create_prob_flag = True
          End Try
          If Not rk Is Nothing Then
            rk.Close()
          End If
      End If
    End If
  End If
End Sub
Private Sub Delete_Logo_Value(ByRef logo_dlet_keyname As String, ByRef spredefinedkey As Integer, ByRef dletvalue As String)
  If spredefinedkey = HKEY_LOCAL_MACHINE Then
    Try
      rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_dlet_keyname, True)
    Catch
      Err.Clear()
      create_prob_flag = True
    End Try
    If Not rk Is Nothing Then
       Try
         rk.DeleteValue(dletvalue, False)
       Catch
         Err.Clear()
         create_prob_flag = True
       End Try
       If Not rk Is Nothing Then
         rk.Close()
       End If
    End If
  Else
    Try
      rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_dlet_keyname, True)
    Catch
      Err.Clear()
      create_prob_flag = True
    End Try
    If Not rk Is Nothing Then
      Try
        rk.DeleteValue(dletvalue, False)
      Catch
        Err.Clear()
        create_prob_flag = True
      End Try
      If Not rk Is Nothing Then
        rk.Close()
      End If
    End If
  End If
End Sub
Private Function getie7cmdmapnum(ByRef ie7_cmd_parm_value As String) As Integer
  logo_sub_keyname = "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
  logo_sub_valuename = ie7_cmd_parm_value
  logo_lreturnvalue = 0
  Try
    rk3 = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
  Catch
    Err.Clear()
    logo_lreturnvalue = 2
  End Try
  If rk3 Is Nothing Then
      logo_lreturnvalue = 2
  Else
      logo_lreturnvalue = 0
  End If
  logo_string_ret_length = 8
  Dim logo_string_ret_length_2 As Integer
  logo_string_ret_length_2 = 8
  If logo_lreturnvalue = 0 Then
   If Me.win2kxp_flag = "Y" Then
     Try
       ie7_value_returned = rk3.GetValue(logo_sub_valuename, "notfound")
     Catch
       Err.Clear()
       logo_lreturnvalue = 2
     End Try
     If ie7_value_returned.ToString = "notfound" Then
         logo_lreturnvalue = 2
     Else
       If logo_lreturnvalue = 0 Then
         If ie7_value_returned.GetType.ToString = "System.Int32" Or _
            ie7_value_returned.GetType.ToString = "System.String" Then
            logo_string_returned = ie7_value_returned.ToString.Length
            If logo_string_returned >= 1 Then
              If IsNumeric(ie7_value_returned.ToString) Then
                logo_string_returned = ie7_value_returned.ToString
                logo_long_byte = Mid(logo_string_returned, 1, 4)
                logo_long_returned = CInt(logo_long_byte)
                getie7cmdmapnum = logo_long_returned
                logo_lreturnvalue = 0
              Else
                getie7cmdmapnum = 0
              End If
            Else
              getie7cmdmapnum = 0
            End If
         Else
            getie7cmdmapnum = 0
         End If
       End If
     End If
   Else
     getie7cmdmapnum = 0
   End If
  Else
    getie7cmdmapnum = 0
  End If
  If Not rk3 Is Nothing Then
      rk3.Close()
  End If
End Function
Private Function setie7cmdmapnum(ByRef ie7_cmd_parm_value_set As String, ByRef ie7_cmd_parm_double As Integer) As Integer
  logo_lreturnvalue = 0
  logo_subkey = "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
  logo_string_key_value = ie7_cmd_parm_value_set
  Try
    rk2 = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_subkey, True)
  Catch
    Err.Clear()
    logo_lreturnvalue = 2
    create_prob_flag = True
  End Try
  If rk2 Is Nothing Then
      logo_lreturnvalue = 2
  Else
      logo_lreturnvalue = 0
  End If
  If logo_lreturnvalue = 0 Then
      logo_subkey = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
      logo_cmdvalue_obj = ie7_cmd_parm_double
      If Not rk2 Is Nothing Then
        Try
          rk2.SetValue(ie7_cmd_parm_value_set, logo_cmdvalue_obj, RegistryValueKind.DWord)
        Catch
          Err.Clear()
          create_prob_flag = True
        End Try
      End If
      If Not rk2 Is Nothing Then
          rk2.Close()
      End If
  End If
End Function
Private Sub evaluate_IE7_buttons()
  logo_IE7_first_mapid_found_flag = False
  logo_IE7_first_mapid_pos = 0
  ReDim logo_IE7_default_toolbar_map(logo_IE7_string_ret_length + 1)
  For logo_IE7_i = 0 To logo_IE7_string_ret_length - 1
      logo_IE7_default_toolbar_map(logo_IE7_i) = logo_IE7_default_toolbar_byte(logo_IE7_i)
  Next  '02/28/2007 removed logo_IE7_i variable for performance'
  toolbar_IE7_cmd_count = 1
  logo_IE7_current_button_num = 4
  logo_num_of_IE7_separators = 0
  logo_IE7_j = 0
  For logo_IE7_j = 1 To logo_IE7_num_of_toolbar_items
   If (logo_IE7_default_toolbar_byte(logo_IE7_current_button_num) = 255) Or _
    (logo_IE7_default_toolbar_byte(logo_IE7_current_button_num) = 254) Then
    If logo_IE7_default_toolbar_byte(logo_IE7_current_button_num + 1) = 255 Or _
     logo_IE7_default_toolbar_byte(logo_IE7_current_button_num + 1) = 254 Then
      logo_IE7_default_separator = CInt(((logo_IE7_current_button_num - 4) / 4))
      logo_IE7_separator_found = True
      logo_num_of_IE7_separators = logo_num_of_IE7_separators + 1
      If logo_num_of_IE7_separators = 1 Then
          logo_IE7_pos_sep_1 = logo_IE7_default_separator
      End If
      If logo_num_of_IE7_separators = 2 Then
          logo_IE7_pos_sep_2 = logo_IE7_default_separator
      End If
    End If
   End If
      If logo_IE7_default_toolbar_byte(logo_IE7_current_button_num + 1) < 254 Then
          dest_hi_value = logo_IE7_default_toolbar_byte(logo_IE7_current_button_num + 0)
       If IE9_flag = True Then
              If dest_hi_value > cmd_id_cutoff_value_IE9 Then
               If logo_IE7_first_mapid_found_flag = False Then
                   If dest_hi_value <= max_poss_cmd_id Then
                       logo_IE7_first_mapid_pos = logo_IE7_current_button_num
                       logo_IE7_first_mapid_found_flag = True
                   End If
               End If
           End If
       Else
        If IE8_flag = True Then
           If dest_hi_value > cmd_id_cutoff_value_IE8 Then
               If logo_IE7_first_mapid_found_flag = False Then
                   If dest_hi_value <= max_poss_cmd_id Then
                       logo_IE7_first_mapid_pos = logo_IE7_current_button_num
                       logo_IE7_first_mapid_found_flag = True
                   End If
               End If
           End If
         Else
           If dest_hi_value > cmd_id_cutoff_value_IE7 Then
               If logo_IE7_first_mapid_found_flag = False Then
                   If dest_hi_value <= max_poss_cmd_id Then
                       logo_IE7_first_mapid_pos = logo_IE7_current_button_num
                       logo_IE7_first_mapid_found_flag = True
                   End If
               End If
           End If
         End If
       End If
           dest_hi_string = CStr(logo_IE7_default_toolbar_byte(logo_IE7_current_button_num + 0))
           dest_lo_string = CStr(logo_IE7_default_toolbar_byte(logo_IE7_current_button_num + 1))
           If Len(dest_lo_string) < 2 Then
             dest_lo_string = "0" & dest_lo_string
             dest_lo_value = hex2lng2(dest_lo_string)
             dest_hi_value = dest_hi_value + dest_lo_value
             toolbar_string = Hex(dest_hi_value)
           Else
             dest_hi_string = CStr(logo_IE7_default_toolbar_byte(logo_IE7_current_button_num + 0))
             dest_hi_string = Hex(CInt(dest_hi_string))
             dest_lo_string = CStr(logo_IE7_default_toolbar_byte(logo_IE7_current_button_num + 1))
             dest_lo_string = Hex(CInt(dest_lo_string))
           End If
           If Len(dest_hi_string) < 2 Then
               dest_hi_string = "0" & dest_hi_string
           End If
           If Len(dest_lo_string) < 2 Then
               dest_lo_string = "0" & dest_lo_string
           End If
           toolbar_string = dest_lo_string & dest_hi_string
       End If
      logo_IE7_current_button_num = logo_IE7_current_button_num + 4
  Next
  analyze_IE7_buttons()
End Sub
Private Sub analyze_IE7_buttons()
  logo_IE7_bumped_cmdmap_flag = False
  logo_IE7_target_compare_pos = 0
  If logo_IE7_button_installed = True Then
      success_install_msg()
      Exit Sub
  End If
  If logo_IE7_use_mapid_logic = True Then
      If logo_IE7_first_mapid_pos > 0 Then
          If logo_IE7_default_target_pos > 0 Then
              logo_IE7_target_compare_pos = logo_IE7_default_target_pos * 4
          End If
          If logo_IE7_first_mapid_pos < logo_IE7_target_compare_pos Then
              If (logo_IE7_first_mapid_pos Mod 4) = 0 Then
                  logo_IE7_target_button_pos = CInt((logo_IE7_first_mapid_pos / 4))
                  logo_IE7_bumped_cmdmap_flag = True
              Else
                  logo_IE7_target_button_pos = logo_IE7_default_target_pos
              End If
              insert_new_IE7_button()
              success_install_msg()
              Exit Sub
          Else
              If logo_IE7_first_mapid_pos = logo_IE7_target_compare_pos Then
                  If logo_IE7_first_mapid_pos > 0 Then
                      If (logo_IE7_first_mapid_pos Mod 4) = 0 Then
                          logo_IE7_target_button_pos = CInt((logo_IE7_first_mapid_pos / 4))
                          logo_IE7_bumped_cmdmap_flag = True
                          insert_new_IE7_button()
                          success_install_msg()
                          Exit Sub
                      End If
                  End If
              Else
                  logo_IE7_target_button_pos = logo_IE7_default_target_pos
                  insert_new_IE7_button()
                  success_install_msg()
                  Exit Sub
              End If
          End If
      Else
          logo_IE7_target_button_pos = logo_IE7_default_target_pos
          insert_new_IE7_button()
          success_install_msg()
          Exit Sub
      End If
  Else
      logo_IE7_target_button_pos = logo_IE7_default_target_pos
      insert_new_IE7_button()
      success_install_msg()
      Exit Sub
  End If
End Sub
Private Sub insert_new_IE7_button()
  ReDim logo_IE7_default_toolbar_byte(logo_IE7_string_ret_length + 4)
  If logo_IE7_target_button_pos > 1 Then
      If logo_IE7_num_of_toolbar_items > logo_IE7_default_target_pos Then
          logo_IE7_button_difference = logo_IE7_num_of_toolbar_items - logo_IE7_target_button_pos
          If logo_IE7_bumped_cmdmap_flag = False Then
              logo_IE7_num_bytes_to_move = logo_IE7_button_difference * 4
              logo_IE7_start_pos = (logo_IE7_target_button_pos * 4) + 4
          Else
              logo_IE7_num_bytes_to_move = (logo_IE7_button_difference * 4) + 4
              logo_IE7_start_pos = (logo_IE7_target_button_pos * 4)
          End If
      Else
          If logo_IE7_num_of_toolbar_items < logo_IE7_default_target_pos Then
              logo_IE7_target_button_pos = logo_IE7_num_of_toolbar_items
              logo_IE7_button_difference = logo_IE7_num_of_toolbar_items
              If logo_IE7_bumped_cmdmap_flag = False Then
                  logo_IE7_num_bytes_to_move = logo_IE7_button_difference * 4
                  logo_IE7_start_pos = (logo_IE7_target_button_pos * 4) + 4
              Else
                  logo_IE7_num_bytes_to_move = (logo_IE7_button_difference * 4) + 4
                  logo_IE7_start_pos = (logo_IE7_target_button_pos * 4)
              End If
          End If
      End If
  Else
      If logo_IE7_target_button_pos = 1 Then
          logo_IE7_button_difference = logo_IE7_num_of_toolbar_items - 1
          If logo_IE7_bumped_cmdmap_flag = False Then
              logo_IE7_num_bytes_to_move = logo_IE7_button_difference * 4
              logo_IE7_start_pos = (logo_IE7_target_button_pos * 4) + 4
          Else
              logo_IE7_num_bytes_to_move = (logo_IE7_button_difference * 4) + 4
              logo_IE7_start_pos = (logo_IE7_target_button_pos * 4)
          End If
      End If
      If logo_IE7_target_button_pos = 0 Then
          logo_IE7_button_difference = logo_IE7_num_of_toolbar_items
          logo_IE7_num_bytes_to_move = logo_IE7_button_difference * 4
          logo_IE7_start_pos = (logo_IE7_target_button_pos * 4) + 4
      End If
  End If
  If logo_IE7_start_pos >= logo_IE7_string_ret_length Then
      For logo_IE7_i = 0 To logo_IE7_string_ret_length
          logo_IE7_default_toolbar_byte(logo_IE7_i) = logo_IE7_default_toolbar_map(logo_IE7_i)
      Next
  Else
      For logo_IE7_i = 0 To logo_IE7_string_ret_length
          logo_IE7_default_toolbar_byte(logo_IE7_i) = logo_IE7_default_toolbar_map(logo_IE7_i)
      Next
  End If
  If logo_IE7_start_pos >= logo_IE7_string_ret_length Then
  Else
      For logo_IE7_i = logo_IE7_start_pos To ((logo_IE7_start_pos + logo_IE7_num_bytes_to_move) - 1)
          logo_IE7_default_toolbar_byte(logo_IE7_i + 4) = logo_IE7_default_toolbar_map(logo_IE7_i)
      Next
  End If
  If logo_IE7_bumped_cmdmap_flag = False Then
      logo_IE7_start_byte_pos = ((logo_IE7_target_button_pos * 4) + 4)
  Else
      logo_IE7_start_byte_pos = ((logo_IE7_target_button_pos * 4))
  End If
    If IE8_flag = True Or IE9_flag = True Then
        If IE7_cmdcount > 0 Then
            If IE7_current_cmd_id < 105 Then
                IE7_cmdcount = 105
                IE7_current_cmd_id = IE7_cmdcount
            End If
        Else
            IE7_cmdcount = 105
            IE7_current_cmd_id = IE7_cmdcount
        End If
    Else
        If IE7_cmdcount > 0 Then
            If IE7_current_cmd_id < 16 Then
                IE7_cmdcount = 16
                IE7_current_cmd_id = IE7_cmdcount
            End If
        Else
            IE7_cmdcount = 16
            IE7_current_cmd_id = IE7_cmdcount
        End If
    End If
  logo_IE7_default_toolbar_byte(logo_IE7_start_byte_pos) = CByte(IE7_current_cmd_id)
  logo_IE7_default_toolbar_byte(logo_IE7_start_byte_pos + 1) = &H1S
  logo_IE7_default_toolbar_byte(logo_IE7_start_byte_pos + 2) = &H0S
  logo_IE7_default_toolbar_byte(logo_IE7_start_byte_pos + 3) = &H0S
  logo_default_IE7_toolbar_length = logo_IE7_string_ret_length
  logo_default_IE7_toolbar_length = logo_default_IE7_toolbar_length + 4
  logo_IE7_subkey = "Software\Microsoft\Internet Explorer\LowRegistry\CommandBar"
  If IE8_flag = True Or IE9_flag = True Then
      logo_IE7_string_key_value = "CommandBandLayout"
  Else
      logo_IE7_string_key_value = "ButtonLayout"
  End If
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
  Catch
    Err.Clear()
    create_prob_flag = True
  End Try
  If rk Is Nothing Then
      logo_IE7_lreturnvalue = 2
  Else
      logo_IE7_lreturnvalue = 0
  End If
  If logo_IE7_lreturnvalue = 0 Then
    Try
      rk.SetValue(logo_IE7_string_key_value, logo_IE7_default_toolbar_byte, RegistryValueKind.Binary)
    Catch
      Err.Clear()
      create_prob_flag = True
    End Try
  End If
  If logo_IE7_lreturnvalue = 0 Then
      logo_IE7_button_installed = True
  Else
      logo_IE7_button_installed = False
  End If
  If Not rk Is Nothing Then
      rk.Close()
  End If
End Sub
Private Sub match_b4_n_after()
  For match_b4_loop = 0 To IE7_b4_cmdcount_hklm_both - 1
     IE7_search_cmd_id = cmd_IE7_b4_array(match_b4_loop).ce7_b4_cmdvalue
      If IE7_search_cmd_id > 0 Then
          check_for_cmdid_install()
          If logo_b4_toolbar_prev_install = True Then
              cmd_IE7_b4_array(match_b4_loop).ce7_b4_mapped_flag = True
          Else
              cmd_IE7_b4_array(match_b4_loop).ce7_b4_mapped_flag = False
          End If
      Else
          cmd_IE7_b4_array(match_b4_loop).ce7_b4_mapped_flag = False
      End If
  Next
  For match_b4_loop = 0 To IE7_b4_cmdcount_hklm_both - 1
      IE7_search_cmd_id = cmd_IE7_b4_array(match_b4_loop).ce7_b4_cmdvalue
      If cmd_IE7_b4_array(match_b4_loop).ce7_b4_mapped_flag = True Then
          For match_ie_loop = 0 To IE7_cmdcount_hklm_both - 1
              If cmd_IE7_b4_array(match_b4_loop).ce7_valuename = cmd_IE7_array(match_ie_loop).ce_valuename Then
                  cmd_IE7_b4_array(match_b4_loop).ce7_IE7_cmdvalue = cmd_IE7_array(match_ie_loop).ce_IE7_cmdvalue
                  cmd_IE7_array(match_ie_loop).ce_b4_cmdvalue = cmd_IE7_b4_array(match_b4_loop).ce7_b4_cmdvalue
                  Exit For
              End If
          Next
          If cmd_IE7_b4_array(match_b4_loop).ce7_IE7_cmdvalue = cmd_IE7_b4_array(match_b4_loop).ce7_b4_cmdvalue Then
          'continue'
          Else
              replace_b4_n_after()
          End If
      End If
  Next
End Sub
Private Sub match_for_delete()
  For match_b4_loop = 0 To IE7_cmdcount_hklm_both - 1
      IE7_search_cmd_id = 0
      For match_ie_loop = 0 To IE7_b4_cmdcount_hklm_both - 1
          If cmd_IE7_array(match_b4_loop).ce_valuename = cmd_IE7_b4_array(match_ie_loop).ce7_valuename Then
              cmd_IE7_array(match_b4_loop).ce_b4_cmdvalue = cmd_IE7_b4_array(match_ie_loop).ce7_b4_cmdvalue
              IE7_search_cmd_id = cmd_IE7_b4_array(match_ie_loop).ce7_b4_cmdvalue
              Exit For
          End If
      Next
      If IE7_search_cmd_id > 0 Then
          check_for_cmdid_delete()
          If logo_b4_toolbar_prev_install = True Then
              cmd_IE7_array(match_b4_loop).ce_mapped_flag = True
          Else
              cmd_IE7_array(match_b4_loop).ce_mapped_flag = False
          End If
      Else
          cmd_IE7_array(match_b4_loop).ce_mapped_flag = False
      End If
  Next
  For match_b4_loop = 0 To IE7_cmdcount_hklm_both
      IE7_search_cmd_id = cmd_IE7_array(match_b4_loop).ce_b4_cmdvalue
      If cmd_IE7_array(match_b4_loop).ce_IE7_cmdvalue = cmd_IE7_array(match_b4_loop).ce_b4_cmdvalue Then
          'continue'
      Else
          replace_after_delete()
      End If
  Next
  update_new_IE7_button()
End Sub
Private Sub replace_after_delete()
  Dim rep_b4_cmd_id_value As Integer
  Dim rep_ie7_cmd_id_value As Integer
  Dim rep_logo_d As Integer
  If IE9_flag = True Then
      cmd_id_compare_value = cmd_id_compare_value_IE9
  Else
    If IE8_flag = True Then
        cmd_id_compare_value = cmd_id_compare_value_IE8
    Else
        cmd_id_compare_value = cmd_id_compare_value_IE7
    End If
  End If
  If cmd_IE7_array(match_b4_loop).ce_IE7_cmdvalue > cmd_id_compare_value Then
      If cmd_IE7_array(match_b4_loop).ce_b4_cmdvalue > cmd_id_compare_value Then
          rep_b4_cmd_id_value = cmd_IE7_array(match_b4_loop).ce_b4_cmdvalue
          rep_ie7_cmd_id_value = cmd_IE7_array(match_b4_loop).ce_IE7_cmdvalue
          If logo_IE7_string_ret_length > 4 Then
              logo_IE7_begin_button_num = logo_IE7_default_toolbar_byte(4)
              logo_IE7_num_of_toolbar_items = CInt(((logo_IE7_string_ret_length - 4) / 4))
              For rep_logo_d = 0 To logo_IE7_num_of_toolbar_items - 1
                  logo_IE7_last_button_seq = (rep_logo_d * 4) + 4
                  logo_IE7_last_button_value = logo_IE7_default_toolbar_byte(logo_IE7_last_button_seq)
                  If rep_b4_cmd_id_value = logo_IE7_last_button_value Then
                      If logo_IE7_last_button_seq = cmd_IE7_b4_array(match_b4_loop).ce7_b4_inst_pos Then
                          logo_IE7_default_toolbar_byte(logo_IE7_last_button_seq) = CByte(rep_ie7_cmd_id_value)
                          num_of_IE7_logo_buttons = num_of_IE7_logo_buttons + 1
                      End If
                  End If
              Next
          End If
      End If
  End If
End Sub
Private Sub replace_b4_n_after()
  Dim rep_b4_cmd_id_value As Integer
  Dim rep_ie7_cmd_id_value As Integer
  Dim rep_logo_d As Integer
  If IE9_flag = True Then                                 'IE9 upgrade - 09/17/2010'
    cmd_id_compare_value = cmd_id_compare_value_IE9
  Else
    If IE8_flag = True Then
        cmd_id_compare_value = cmd_id_compare_value_IE8
    Else
        cmd_id_compare_value = cmd_id_compare_value_IE7
    End If
  End If
  If cmd_IE7_b4_array(match_b4_loop).ce7_IE7_cmdvalue > cmd_id_compare_value Then
      If cmd_IE7_b4_array(match_b4_loop).ce7_b4_cmdvalue > cmd_id_compare_value Then
          rep_b4_cmd_id_value = cmd_IE7_b4_array(match_b4_loop).ce7_b4_cmdvalue
          rep_ie7_cmd_id_value = cmd_IE7_b4_array(match_b4_loop).ce7_IE7_cmdvalue
          If logo_IE7_string_ret_length > 4 Then
              logo_IE7_begin_button_num = logo_IE7_default_toolbar_byte(4)
              logo_IE7_num_of_toolbar_items = CInt(((logo_IE7_string_ret_length - 4) / 4))
              For rep_logo_d = 0 To logo_IE7_num_of_toolbar_items - 1
                  logo_IE7_last_button_seq = (rep_logo_d * 4) + 4
                  logo_IE7_last_button_value = logo_IE7_default_toolbar_byte(logo_IE7_last_button_seq)
                  If rep_b4_cmd_id_value = logo_IE7_last_button_value Then
                      If logo_IE7_last_button_seq = cmd_IE7_b4_array(match_b4_loop).ce7_b4_inst_pos Then
                          logo_IE7_default_toolbar_byte(logo_IE7_last_button_seq) = CByte(rep_ie7_cmd_id_value)
                          num_of_IE7_logo_buttons = num_of_IE7_logo_buttons + 1
                      End If
                  End If
              Next
          End If
      End If
  End If
End Sub
Private Sub match_b4_cmdmaps()
  logo_b4_cmdid_installed = False
  For logo_b4_i = 0 To IE7_b4_cmdcount_hklm_both - 1
      If Mid(logo_new_guid_2, 1, 38) = Mid(cmd_IE7_b4_array(logo_b4_i).ce7_valuename, 1, 38) Then
          b4_current_cmd_id = cmd_IE7_b4_array(logo_b4_i).ce7_b4_cmdvalue
          logo_b4_cmdid_installed = True
      End If
  Next
End Sub
Private Sub match_IE7_cmdmaps()
  logo_IE7_cmdid_installed = False
  For logo_IE7_i = 0 To IE7_cmdcount_hklm_both - 1
      If Mid(logo_new_guid_2, 1, 38) = Mid(cmd_IE7_array(logo_IE7_i).ce_valuename, 1, 38) Then
          IE7_current_cmd_id = cmd_IE7_array(logo_IE7_i).ce_IE7_cmdvalue
          logo_IE7_cmdid_installed = True
          Exit For
      End If
  Next
End Sub
Private Sub match_map_ids()
  Dim current_map_id As String
  Dim map_id_counter As Integer
  logo_mapid_found_flag = False
  logo_first_mapid_pos = 0
  If toolbar_cmd_count > 1 Then
      For logo_i = 1 To toolbar_cmd_count
          current_map_id = toolbar_cmd_array(logo_i).tm_valuehex
          For map_id_counter = 1 To cmdcount
              If Mid(cmdarray(map_id_counter).ce_valuehex, 1, 4) = Mid(current_map_id, 1, 4) Then
                  If Len(current_map_id) > 3 Then
                      logo_first_mapid_pos = logo_i - 1
                      logo_first_matched_id = current_map_id
                      logo_mapid_found_flag = True
                      Exit For
                  End If
              End If
          Next
          If logo_first_mapid_pos > 0 Or logo_mapid_found_flag = True Then
              Exit For
          End If
      Next
  End If
End Sub
Private Sub build_IE7_cmdmapping()
  Dim logo_nextid_changed As Boolean
  logo_nextid_changed = False
  logo_sub_keyname = "Software\Microsoft\Internet Explorer\Extensions"
  logo_sub_valuename = ""
  logo_lreturnvalue = 0
  Try
    rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_sub_keyname, False)
  Catch
    Err.Clear()
    logo_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_lreturnvalue = 2
  Else
      logo_lreturnvalue = 0
  End If
  If logo_lreturnvalue = 2 Then
    netnewpath = "Software\Microsoft\Internet Explorer\Extensions"
    If Not create_rk Is Nothing Then
      create_rk.Close()
    End If
    Try
      create_rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(netnewpath)
    Catch
      Err.Clear()
    End Try
    If create_rk Is Nothing Then
      logo_create_rc = 2
      create_prob_flag = True
    Else
      logo_create_rc = 0
    End If
    Try
      rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_sub_keyname, False)
    Catch
      Err.Clear()
    End Try
  End If
  If rk Is Nothing Then
      logo_lreturnvalue = 2
      logo_rqve_lpcsubkeys = 0
  Else
      logo_lreturnvalue = 0
      logo_rqve_lpcsubkeys = rk.SubKeyCount
  End If
  Dim num_of_enum_keys As Integer
  num_of_enum_keys = logo_rqve_lpcsubkeys
  For key_dwindex = 0 To 90
      cmd_ext_array(key_dwindex).ea_extension_name = Nothing
  Next
  If logo_rqve_lpcsubkeys > 0 Then
    key_dwindex = 0
    key_strvaluelen = 255
    key_strvalue = Space(254) & Chr(0)
    Try
      rk.GetSubKeyNames()
    Catch
      Err.Clear()
      create_prob_flag = True
    End Try
    For Each subKeyName As String In rk.GetSubKeyNames()
        If Mid(subKeyName.ToString, 1, 1) = "{" Then
            cmd_ext_array(key_dwindex).ea_extension_name = subKeyName.ToString
            If key_dwindex < logo_rqve_lpcsubkeys - 1 Then
                key_dwindex = key_dwindex + 1
            End If
        End If
    Next
    If Not rk Is Nothing Then
        rk.Close()
    End If
    logo_sub_keyname = "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
    logo_sub_valuename = ""
    Try
      rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, True)
    Catch
      Err.Clear()
      logo_lreturnvalue = 2
    End Try
    If rk Is Nothing Then
        logo_lreturnvalue = 2
    Else
        logo_lreturnvalue = 0
    End If
    If logo_lreturnvalue = 2 Then
      netnewpath = logo_sub_keyname
      If Not create_rk Is Nothing Then
        create_rk.Close()
      End If
      Try
        create_rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(netnewpath)
      Catch
        Err.Clear()
      End Try
      If create_rk Is Nothing Then
        logo_create_rc = 2
        create_prob_flag = True
      Else
        logo_create_rc = 0
      End If
      Try
        rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
      Catch
        Err.Clear()
      End Try
      If rk Is Nothing Then
        logo_lreturnvalue = 2
      Else
        logo_lreturnvalue = 0
      End If
    End If
    key_dwindex = 0
    For key_dwindex = 0 To logo_rqve_lpcsubkeys - 1
        logo_sub_valuename = cmd_ext_array(key_dwindex).ea_extension_name
        If Mid(logo_sub_valuename, 1, 1) = "{" Then
          Try
            ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
          Catch
            Err.Clear()
          End Try
            If ie7_value_returned.ToString = "notfound" Then
                logo_lreturnvalue = 2
                logo_cmdvalue_obj = IE7_logo_next_avail_id
                logo_sub_keyname = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
              If Not rk Is Nothing Then
                Try
                 rk.SetValue(logo_sub_valuename, logo_cmdvalue_obj, RegistryValueKind.DWord)
                Catch
                 Err.Clear()
                 create_prob_flag = True
                End Try
              End If
                IE7_logo_next_avail_id = IE7_logo_next_avail_id + 1
                logo_nextid_changed = True
            Else
                logo_lreturnvalue = 0
            End If
            If Mid(logo_new_guid, 1, 38) = Mid(logo_sub_valuename, 1, 38) Then
                logo_IE7_install_guid = "Y"
                logo_IE7_install_index = rev_dwindex
            End If
            If Mid(logo_new_guid_2, 1, 38) = Mid(logo_sub_valuename, 1, 38) Then
                logo_IE7_install_guid_2 = "Y"
                logo_IE7_install_index_2 = rev_dwindex
            End If
        End If
    Next
    If logo_nextid_changed = True Then
        logo_return_value = setie7cmdmapnum("NextId", IE7_logo_next_avail_id)
        If Not rk Is Nothing Then
            rk.Close()
        End If
    End If
  End If
  If Not rk Is Nothing Then
      rk.Close()
  End If
  logo_nextid_changed = False
  logo_sub_keyname = "Software\Microsoft\Internet Explorer\Extensions"
  logo_sub_valuename = ""
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
  Catch
    Err.Clear()
    create_prob_flag = True
  End Try
  If rk Is Nothing Then
      logo_lreturnvalue = 2
      logo_rqve_lpcsubkeys = 0
  Else
      logo_lreturnvalue = 0
      logo_rqve_lpcsubkeys = rk.SubKeyCount
  End If
  If logo_lreturnvalue = 2 Then
    netnewpath = "Software\Microsoft\Internet Explorer\Extensions"
    If Not create_rk Is Nothing Then
      create_rk.Close()
    End If
    Try
      create_rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(netnewpath)
    Catch
      Err.Clear()
      logo_lreturnvalue = 2
    End Try
    Try
      rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
    Catch
      Err.Clear()
      create_prob_flag = True
    End Try
  End If
  If rk Is Nothing Then
    logo_lreturnvalue = 2
    logo_rqve_lpcsubkeys = 0
  Else
    logo_lreturnvalue = 0
    logo_rqve_lpcsubkeys = rk.SubKeyCount
  End If
  If logo_rqve_lpcsubkeys > 0 Then
      num_of_enum_keys = logo_rqve_lpcsubkeys
      For key_dwindex = 0 To 90
          cmd_ext_array(key_dwindex).ea_extension_name = Nothing
      Next
      If logo_rqve_lpcsubkeys > 0 Then
          key_dwindex = 0
          key_strvaluelen = 255
          key_strvalue = Space(254) & Chr(0)
          Try
            rk.GetSubKeyNames()
          Catch
            Err.Clear()
            create_prob_flag = True
          End Try
          For Each subKeyName As String In rk.GetSubKeyNames()
              'Dim tempKey As RegistryKey = _
              '    rk.OpenSubKey(subKeyName)  - fixed bug of not processing hkcu extensions 11/26/2013'
             If Mid(subKeyName.ToString, 1, 1) = "{" Then
                cmd_ext_array(key_dwindex).ea_extension_name = subKeyName.ToString
               If key_dwindex < logo_rqve_lpcsubkeys - 1 Then
                   key_dwindex = key_dwindex + 1
               End If
             End If
             ' If key_dwindex < logo_rqve_lpcsubkeys - 1 Then
             '     key_dwindex = key_dwindex + 1
             ' End If
          Next
          If Not rk Is Nothing Then
              rk.Close()
          End If
      End If
      logo_sub_keyname = "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
      logo_sub_valuename = ""
      Try
        rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, True)
      Catch
        Err.Clear()
        create_prob_flag = True
        logo_lreturnvalue = 2
      End Try
      For key_dwindex = 0 To logo_rqve_lpcsubkeys - 1
        logo_sub_valuename = cmd_ext_array(key_dwindex).ea_extension_name
        If Mid(logo_sub_valuename, 1, 1) = "{" Then
          ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
          If ie7_value_returned.ToString = "notfound" Then
              logo_lreturnvalue = 2
          Else
              logo_lreturnvalue = 0
          End If
          If logo_lreturnvalue = 2 Then
              logo_sub_keyname = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
              logo_cmdvalue_obj = IE7_logo_next_avail_id
             If Not rk Is Nothing Then
                Try
                  rk.SetValue(logo_sub_valuename, logo_cmdvalue_obj, RegistryValueKind.DWord)
                Catch
                  Err.Clear()
                  create_prob_flag = True
                End Try
             End If
              IE7_logo_next_avail_id = IE7_logo_next_avail_id + 1
              logo_nextid_changed = True
          End If
          If Mid(logo_new_guid, 1, 38) = Mid(logo_sub_valuename, 1, 38) Then
              logo_IE7_install_guid = "Y"
              logo_IE7_install_index = rev_dwindex
          End If
          If Mid(logo_new_guid_2, 1, 38) = Mid(logo_sub_valuename, 1, 38) Then
              logo_IE7_install_guid_2 = "Y"
              logo_IE7_install_index_2 = rev_dwindex
          End If
         End If
      Next
      If logo_nextid_changed = True Then
          logo_return_value = setie7cmdmapnum("NextId", IE7_logo_next_avail_id)
          If Not rk Is Nothing Then
              rk.Close()
          End If
      End If
  End If
End Sub
Private Sub build_IE7_default_toolbar()
  ReDim logo_IE7_default_toolbar_map(27)
  ReDim logo_IE7_default_toolbar_byte(27)
  If IE8_flag = False Then
      logo_IE7_default_toolbar_map(0) = &H7
      logo_IE7_default_toolbar_map(1) = &H0
      logo_IE7_default_toolbar_map(2) = &H0
      logo_IE7_default_toolbar_map(3) = &H0
      logo_IE7_default_toolbar_map(4) = &H1
      logo_IE7_default_toolbar_map(5) = &H1
      logo_IE7_default_toolbar_map(6) = &H0
      logo_IE7_default_toolbar_map(7) = &H0
      logo_IE7_default_toolbar_map(8) = &H3
      logo_IE7_default_toolbar_map(9) = &H1
      logo_IE7_default_toolbar_map(10) = &H0
      logo_IE7_default_toolbar_map(11) = &H0
      logo_IE7_default_toolbar_map(12) = &H4
      logo_IE7_default_toolbar_map(13) = &H1
      logo_IE7_default_toolbar_map(14) = &H0
      logo_IE7_default_toolbar_map(15) = &H0
      logo_IE7_default_toolbar_map(16) = &H5
      logo_IE7_default_toolbar_map(17) = &H1
      logo_IE7_default_toolbar_map(18) = &H0
      logo_IE7_default_toolbar_map(19) = &H0
      logo_IE7_default_toolbar_map(20) = &H6
      logo_IE7_default_toolbar_map(21) = &H1
      logo_IE7_default_toolbar_map(22) = &H0
      logo_IE7_default_toolbar_map(23) = &H0
      logo_IE7_default_toolbar_map(24) = &H7
      logo_IE7_default_toolbar_map(25) = &H1
      logo_IE7_default_toolbar_map(26) = &H0
      logo_IE7_default_toolbar_map(27) = &H0
  Else
      logo_IE7_default_toolbar_map(0) = &H7                'IE8 Upgrade'
      logo_IE7_default_toolbar_map(1) = &H0
      logo_IE7_default_toolbar_map(2) = &H0
      logo_IE7_default_toolbar_map(3) = &H0
      logo_IE7_default_toolbar_map(4) = &H1                'IE8 Upgrade'
      logo_IE7_default_toolbar_map(5) = &H1                'IE8 Upgrade'
      logo_IE7_default_toolbar_map(6) = &H0                'IE8 Upgrade'
      logo_IE7_default_toolbar_map(7) = &H0                'IE8 Upgrade'
      logo_IE7_default_toolbar_map(8) = &H3                'IE8 Upgrade'
      logo_IE7_default_toolbar_map(9) = &H1                'IE8 Upgrade'
      logo_IE7_default_toolbar_map(10) = &H0               'IE8 Upgrade'
      logo_IE7_default_toolbar_map(11) = &H0               'IE8 Upgrade'      
      logo_IE7_default_toolbar_map(12) = &H4               'IE8 Upgrade'
      logo_IE7_default_toolbar_map(13) = &H1               'IE8 Upgrade'
      logo_IE7_default_toolbar_map(14) = &H0               'IE8 Upgrade'
      logo_IE7_default_toolbar_map(15) = &H0               'IE8 Upgrade'      
      logo_IE7_default_toolbar_map(16) = &H5               'IE8 Upgrade'
      logo_IE7_default_toolbar_map(17) = &H1               'IE8 Upgrade'
      logo_IE7_default_toolbar_map(18) = &H0               'IE8 Upgrade'
      logo_IE7_default_toolbar_map(19) = &H0               'IE8 Upgrade'      
      logo_IE7_default_toolbar_map(20) = &H6               'IE8 Upgrade'
      logo_IE7_default_toolbar_map(21) = &H1               'IE8 Upgrade'
      logo_IE7_default_toolbar_map(22) = &H0               'IE8 Upgrade'
      logo_IE7_default_toolbar_map(23) = &H0               'IE8 Upgrade'      
      logo_IE7_default_toolbar_map(24) = &H7               'IE8 Upgrade'
      logo_IE7_default_toolbar_map(25) = &H1               'IE8 Upgrade'
      logo_IE7_default_toolbar_map(26) = &H0               'IE8 Upgrade'
      logo_IE7_default_toolbar_map(27) = &H0               'IE8 Upgrade'      
  End If                                                   'IE8 Upgrade'
  logo_IE7_default_toolbar_length = 27
  logo_IE7_string_ret_length = 27
  For logo_IE7_i = 0 To logo_IE7_default_toolbar_length
      logo_IE7_default_toolbar_byte(logo_IE7_i) = logo_IE7_default_toolbar_map(logo_IE7_i)
  Next
  logo_IE7_subkey = "Software\Microsoft\Internet Explorer\LowRegistry\CommandBar"
  If IE9_flag = True Then                                 'IE9 Upgrade'
    logo_IE7_string_key_value = "CommandBandLayout"
  Else
    If IE8_flag = True Then
        logo_IE7_string_key_value = "CommandBandLayout"
    Else
        logo_IE7_string_key_value = "ButtonLayout"
    End If
  End If
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
  Catch
    Err.Clear()
    create_prob_flag = True
  End Try
  If rk Is Nothing Then
      netnewpath = logo_IE7_subkey
      If Not create_rk Is Nothing Then
        create_rk.Close()
      End If
      Try
        create_rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(netnewpath)
      Catch
        Err.Clear()
      End Try
      If create_rk Is Nothing Then
        logo_create_rc = 2
        create_prob_flag = True
      Else
        logo_create_rc = 0
      End If
      Try
        rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
      Catch
        Err.Clear()
        create_prob_flag = True
      End Try
  End If
  logo_cmdvalue_obj = logo_IE7_default_toolbar_byte(logo_IE7_default_toolbar_length)
  Try
    rk.SetValue(logo_IE7_string_key_value, logo_IE7_default_toolbar_byte, RegistryValueKind.Binary)
  Catch
    Err.Clear()
    create_prob_flag = True
  End Try
  If Not rk Is Nothing Then
      rk.Close()
  End If
  logo_IE7_default_toolbar_built = True
End Sub
Private Sub check_cmdbar_enabled()
  logo_lreturnvalue = 0
  logo_IE7_subkey = "Software\Microsoft\Internet Explorer\CommandBar"
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
  Catch
    Err.Clear()
  End Try
  If rk Is Nothing Then
      logo_lreturnvalue = 2
  Else
      logo_lreturnvalue = 0
  End If
  If logo_lreturnvalue = 2 Or logo_lreturnvalue = 6 Then
      If Not create_rk Is Nothing Then
        create_rk.Close()
      End If
      netnewpath = "Software\Microsoft\Internet Explorer\CommandBar"
      Try
        create_rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(netnewpath)
      Catch
        Err.Clear()
        logo_lreturnvalue = 2
      End Try
      If create_rk Is Nothing Then
        logo_lreturnvalue = 2
        create_prob_flag = True
      Else
        logo_lreturnvalue = 0
      End If
      Try
        rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
      Catch
        Err.Clear()
        create_prob_flag = True
      End Try
      logo_IE7_string_key_value = "CommandBarEnabled"
      If Not rk Is Nothing Then
        Try
          rk.SetValue(logo_IE7_string_key_value, 1, RegistryValueKind.DWord)
        Catch
          Err.Clear()
          logo_lreturnvalue = 2
        End Try
      End If
      If Not rk Is Nothing Then
          rk.Close()
      End If
  Else
      If Not rk Is Nothing Then
          rk.Close()
      End If
      Try
        rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
      Catch
        Err.Clear()
        create_prob_flag = True
      End Try
      logo_IE7_string_key_value = "CommandBarEnabled"
      If Not rk Is Nothing Then
        Try
          rk.SetValue(logo_IE7_string_key_value, 1, RegistryValueKind.DWord)
        Catch
          Err.Clear()
          create_prob_flag = True
        End Try
      End If
      If Not rk Is Nothing Then
          rk.Close()
      End If
  End If
End Sub
Private Sub check_cmdbar_MINIE_enabled()
  logo_lreturnvalue = 0
  logo_IE7_subkey = "Software\Microsoft\Internet Explorer\MINIE"
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
  Catch
    Err.Clear()
  End Try
  If rk Is Nothing Then
      logo_lreturnvalue = 2
  Else
      logo_lreturnvalue = 0
  End If
  If logo_lreturnvalue = 2 Or logo_lreturnvalue = 6 Then
      If Not create_rk Is Nothing Then
        create_rk.Close()
      End If
      netnewpath = "Software\Microsoft\Internet Explorer\MINIE"
      Try
        create_rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(netnewpath)
      Catch
        Err.Clear()
        logo_lreturnvalue = 2
      End Try
      If create_rk Is Nothing Then
        logo_lreturnvalue = 2
        create_prob_flag = True
      Else
        logo_lreturnvalue = 0
      End If
      Try
        rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
      Catch
        Err.Clear()
        create_prob_flag = True
      End Try
      logo_IE7_string_key_value = "CommandBarEnabled"
      If Not rk Is Nothing Then
        Try
          rk.SetValue(logo_IE7_string_key_value, 1, RegistryValueKind.DWord)
        Catch
          Err.Clear()
          logo_lreturnvalue = 2
        End Try
      End If
      If Not rk Is Nothing Then
          rk.Close()
      End If
  Else
      If Not rk Is Nothing Then
          rk.Close()
      End If
      Try
        rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
      Catch
        Err.Clear()
        create_prob_flag = True
      End Try
      logo_IE7_string_key_value = "CommandBarEnabled"
      If Not rk Is Nothing Then
        Try
          rk.SetValue(logo_IE7_string_key_value, 1, RegistryValueKind.DWord)
        Catch
          Err.Clear()
          create_prob_flag = True
        End Try
      End If
      If Not rk Is Nothing Then
          rk.Close()
      End If
  End If
End Sub
Private Sub check_for_b4_toolbar()
  logo_IE7_lreturnvalue = 0
  logo_IE7_subkey = "Software\Microsoft\Internet Explorer\LowRegistry\CommandBar"
  If IE8_flag = True Or IE9_flag = True Then
      logo_IE7_sub_valuename = "CommandBandLayout"
  Else
      logo_IE7_sub_valuename = "ButtonLayout"
  End If
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
  Catch
    Err.Clear()
    logo_IE7_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_IE7_lreturnvalue = 2
  Else
      logo_IE7_lreturnvalue = 0
  End If
  If logo_IE7_lreturnvalue = 0 Then
    Try
      ie7_value_returned = rk.GetValue(logo_IE7_sub_valuename, "notfound")
    Catch
        Err.Clear()
        logo_IE7_lreturnvalue = 2
    End Try
      If ie7_value_returned.ToString = "notfound" Then
          logo_IE7_lreturnvalue = 2
      Else
          If ie7_value_returned.GetType.ToString = "System.Byte[]" Then
           logo_IE7_lreturnvalue = 0
           logo_b4_byte_string = ie7_value_returned
           logo_b4_string_ret_length = logo_b4_byte_string.GetLength(0)
           If logo_b4_string_ret_length > 1 Then
               If logo_b4_string_ret_length > 4 Then
                 logo_IE7_toolbar_remain = (logo_b4_string_ret_length Mod 4)
                 If logo_IE7_toolbar_remain > 0 Then
                   logo_b4_string_ret_length = logo_b4_string_ret_length - logo_IE7_toolbar_remain
                   Array.Resize(logo_b4_byte_string, logo_b4_string_ret_length)
                   Try
                     rk.SetValue(logo_IE7_sub_valuename, logo_b4_byte_string, RegistryValueKind.Binary)
                   Catch
                     create_prob_flag = True
                     Err.Clear()
                   End Try
                 End If
               Else
                  logo_b4_string_ret_length = logo_b4_string_ret_length - 1
               End If
           End If
          End If
      End If
  End If
  If logo_IE7_lreturnvalue = ERROR_FILE_NOT_FOUND Then
      logo_b4_toolbar_cmdmap_exists = "N"
      If Not rk Is Nothing Then
          rk.Close()
      End If
  Else
      If logo_IE7_lreturnvalue = 0 Or logo_IE7_lreturnvalue = ERROR_MORE_DATA Then
          If logo_b4_string_ret_length > 4 Then
              logo_b4_toolbar_cmdmap_exists = "Y"
              ReDim logo_b4_IE7_default_toolbar_byte(logo_b4_string_ret_length + 84)
              logo_b4_IE7_default_toolbar_byte = logo_b4_byte_string
          Else
              logo_b4_toolbar_cmdmap_exists = "N"
          End If
      End If
  End If
End Sub
Private Sub check_for_IE7_cmdmapping()
  logo_lreturnvalue = 0
  logo_sub_keyname = "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
  logo_sub_valuename = ""
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
  Catch
    Err.Clear()
    logo_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_lreturnvalue = 2
  Else
      logo_lreturnvalue = 0
  End If
  If logo_lreturnvalue = 2 Then
      logo_sub_keyname = "Software\Microsoft\Internet Explorer\LowRegistry\Extensions"
      logo_sub_valuename = ""
      Try
        rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
      Catch
        Err.Clear()
        logo_lreturnvalue = 2
        create_prob_flag = True
      End Try
      If rk Is Nothing Then
          logo_lreturnvalue = 2
      Else
          logo_lreturnvalue = 0
      End If
      If logo_lreturnvalue = 0 Then
          netnewpath = "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
          If Not create_rk Is Nothing Then
            create_rk.Close()
          End If
          Try
            create_rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(netnewpath)
          Catch
            create_prob_flag = True
            logo_lreturnvalue = 2
            Err.Clear()
          End Try
          If create_rk Is Nothing Then
            logo_lreturnvalue = 2
            create_prob_flag = True
          Else
            logo_lreturnvalue = 0
          End If
          IE7_logo_next_avail_id = 8192
          logo_return_value = setie7cmdmapnum("NextId", IE7_logo_next_avail_id)
          logo_IE7_build_cmdmap_values = True
      Else
          If logo_lreturnvalue = 2 Then
            netnewpath = "Software\Microsoft\Internet Explorer\LowRegistry\Extensions"
            If Not create_rk Is Nothing Then
              create_rk.Close()
            End If
            Try
              create_rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(netnewpath)
            Catch
              Err.Clear()
              create_prob_flag = True
              logo_lreturnvalue = 2
            End Try
            If create_rk Is Nothing Then
              logo_lreturnvalue = 2
              create_prob_flag = True
            Else
              logo_lreturnvalue = 0
            End If
            netnewpath = "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
            Try
              create_rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(netnewpath)
            Catch
              Err.Clear()
              create_prob_flag = True
              logo_lreturnvalue = 2
            End Try
            IE7_logo_next_avail_id = 8192
            logo_return_value = setie7cmdmapnum("NextId", IE7_logo_next_avail_id)
            logo_IE7_build_cmdmap_values = True
          End If
      End If
      logo_sub_keyname = "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
      Try
        rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, True)
      Catch
        Err.Clear()
        create_prob_flag = True
        logo_lreturnvalue = 2
      End Try
  End If
  logo_lreturnvalue = 0
  If rk Is Nothing Then
    logo_rqve_lpcvalues = 0
    logo_num_cmdmaps = 0
  Else
    logo_rqve_lpcvalues = rk.ValueCount
    logo_num_cmdmaps = logo_rqve_lpcvalues
  End If
  If logo_lreturnvalue = 0 Then
      logo_sub_keyname = "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
      logo_sub_valuename = "NextId"
      Try
        rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, True)
      Catch
        Err.Clear()
        create_prob_flag = True
        logo_lreturnvalue = 2
      End Try
      If rk Is Nothing Then
          logo_lreturnvalue = 2
      Else
          logo_lreturnvalue = 0
      End If
      logo_string_returned = ""
      logo_string_ret_length = 8
      If Me.win2kxp_flag = "Y" Then
          If logo_lreturnvalue = 0 Then
            Try
              ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
            Catch
              Err.Clear()
              logo_lreturnvalue = 2
            End Try
              If ie7_value_returned.ToString = "notfound" Then
                  logo_lreturnvalue = 2
              Else
                If ie7_value_returned.GetType.ToString = "System.Int32" Or _
                   ie7_value_returned.GetType.ToString = "System.String" Then
                   logo_string_ret_length = ie7_value_returned.ToString.Length
                   If logo_string_ret_length >= 1 Then
                     If IsNumeric(ie7_value_returned.ToString) Then
                       logo_long_returned = CInt(ie7_value_returned.ToString)
                       logo_lreturnvalue = 0
                     Else
                       logo_lreturnvalue = 2
                     End If
                   Else
                     logo_lreturnvalue = 2
                   End If
                Else
                  logo_lreturnvalue = 2
                End If
              End If
              If logo_lreturnvalue = 0 Then
                  IE7_logo_next_avail_id = logo_long_returned
                  If IE7_logo_next_avail_id < 8192 Then
                      IE7_logo_next_avail_id = 8192
                      If logo_num_cmdmaps > 1 Then
                          IE7_logo_next_avail_id = IE7_logo_next_avail_id + logo_num_cmdmaps
                      End If
                      logo_cmdvalue_obj = IE7_logo_next_avail_id
                    If Not rk Is Nothing Then
                      Try
                        rk.SetValue(logo_sub_valuename, logo_cmdvalue_obj, RegistryValueKind.DWord)
                      Catch
                        Err.Clear()
                        create_prob_flag = True
                        logo_lreturnvalue = 2
                      End Try
                    End If
                  End If
              End If
          Else
              If logo_lreturnvalue = 2 Then
                  If logo_num_cmdmaps = 0 Then
                      IE7_logo_next_avail_id = 8192
                      logo_return_value = setie7cmdmapnum("NextId", IE7_logo_next_avail_id)
                      logo_IE7_build_cmdmap_values = True
                  End If
                  If logo_num_cmdmaps > 0 Then
                      rev_dwindex = 0
                      rev_hkey = hKey
                      rev_lpvaluename = Space(200)
                      rev_lpcbvaluename = 200
                      IE7_logo_next_avail_id = 8192
                      For rev_dwindex = 0 To logo_num_cmdmaps
                          rev_lpvaluename = Space(200)
                          rev_lpcbvaluename = 200
                          Try
                            rk.SetValue(logo_sub_valuename, IE7_logo_next_avail_id, RegistryValueKind.DWord)
                          Catch
                            Err.Clear()
                            create_prob_flag = True
                          End Try
                      Next
                      logo_return_value = setie7cmdmapnum("NextId", IE7_logo_next_avail_id)
                      logo_IE7_build_cmdmap_values = True
                  End If
              End If
          End If
      End If
  Else
      If logo_lreturnvalue = 2 Then
          netnewpath = "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
          If Not create_rk Is Nothing Then
            create_rk.Close()
          End If
          Try
            create_rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(netnewpath)
          Catch
            Err.Clear()
            create_prob_flag = True
          End Try
          If create_rk Is Nothing Then
            logo_lreturnvalue = 2
            create_prob_flag = True
          Else
            logo_lreturnvalue = 0
          End If
          IE7_logo_next_avail_id = 8192
          logo_return_value = setie7cmdmapnum("NextId", IE7_logo_next_avail_id)
          logo_IE7_build_cmdmap_values = True
      End If
  End If
End Sub
Private Sub check_toolbar_unlocked()     '07/09/2012'
  logo_IE7_lreturnvalue = 0
  logo_IE7_subkey = "Software\Microsoft\Internet Explorer\Main"
  logo_IE7_string_key_value = "Enable Browser Extensions"
  SetUpKeyValue(HKEY_CURRENT_USER, logo_IE7_subkey, "Enable Browser Extensions", "yes")
  logo_IE7_subkey = "Software\Microsoft\Internet Explorer\Toolbar"
  logo_IE7_string_key_value = "Locked"
  logo_IE7_lreturnvalue = 0
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
  Catch
    Err.Clear()
    logo_IE7_lreturnvalue = 2
  End Try
  If logo_IE7_lreturnvalue = 0 Then
    rk.SetValue(logo_IE7_string_key_value, 0, RegistryValueKind.DWord)
  End If
  If Not rk Is Nothing Then
    rk.Close()
  End If
  logo_IE7_lreturnvalue = 0
End Sub

Private Sub check_UAC_setting()
' 11/10/2013 - added for UAC check in Vista, 7, 8'
  Dim logo_string_ret_length_2 As Integer
  Dim logo_check_UAC_value As Long
  logo_sub_keyname = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
  logo_sub_valuename = "EnableLUA"
  logo_lreturnvalue = 0

  Try
    rk3 = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_sub_keyname, False)
  Catch
    Err.Clear()
    logo_lreturnvalue = 2
  End Try
  If rk3 Is Nothing Then
      logo_lreturnvalue = 2
  Else
      logo_lreturnvalue = 0
  End If

  logo_string_ret_length = 8
  logo_string_ret_length_2 = 8
  If logo_lreturnvalue = 0 Then
     Try
       ie7_value_returned = rk3.GetValue(logo_sub_valuename, "notfound")
     Catch
       Err.Clear()
       logo_lreturnvalue = 2
     End Try

     If ie7_value_returned.ToString = "notfound" Then
         logo_lreturnvalue = 2
     Else
       If logo_lreturnvalue = 0 Then
         If ie7_value_returned.GetType.ToString = "System.Int32" Or _
            ie7_value_returned.GetType.ToString = "System.String" Then
            logo_string_returned = ie7_value_returned.ToString.Length
            If logo_string_returned >= 1 Then
              If IsNumeric(ie7_value_returned.ToString) Then
                logo_string_returned = ie7_value_returned.ToString
                logo_long_byte = Mid(logo_string_returned, 1, 4)
                logo_long_returned = CInt(logo_long_byte)
                logo_check_UAC_value = logo_long_returned
                logo_lreturnvalue = 0
              Else
                logo_check_UAC_value = 0
              End If
            Else
              logo_check_UAC_value = 0
            End If
         Else
            logo_check_UAC_value = 0
         End If
       End If
     End If
   Else
     logo_check_UAC_value = 0
   End If

  If Not rk3 Is Nothing Then
      rk3.Close()
  End If
  If logo_check_UAC_value = 1 Then
    win_UAC_enabled_flag = True
  Else
    win_UAC_enabled_flag = False
  End If

'  MsgBox("UAC value:" + logo_check_UAC_value.ToString, MsgBoxStyle.Exclamation, "FirstButton  Installer")'

End Sub

Private Sub correct_IE7_cmdmapping()
  Dim correct_cmd_count As Integer
  logo_lreturnvalue = 0
  correct_cmd_count = 0
  logo_sub_keyname = "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
  logo_sub_valuename = ""
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
  Catch
    Err.Clear()
    logo_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
    logo_lreturnvalue = 2
    logo_rqve_lpcsubkeys = 0
    logo_rqve_lpcvalues = 0
  Else
    logo_lreturnvalue = 0
    logo_rqve_lpcsubkeys = rk.SubKeyCount
    logo_rqve_lpcvalues = rk.ValueCount
  End If
  logo_num_cmdmaps = logo_rqve_lpcvalues
  rev_dwindex = 0
  rev_lpvaluename = ""
  rev_lpcbvaluename = 200
  IE7_logo_next_avail_id = 8192
  If logo_lreturnvalue = 0 Then
   Try
     rk.GetValueNames()
   Catch
     Err.Clear()
     create_prob_flag = True
   End Try
   For Each valueName As String In rk.GetValueNames()
      rev_lpvaluename = ""
      rev_lpcbvaluename = 38
      If Mid(logo_new_guid, 1, 38) = Mid(rev_lpvaluename, 1, 38) Then
          logo_IE7_install_guid = "Y"
          logo_IE7_install_index = correct_cmd_count
      End If
      If Mid(logo_new_guid_2, 1, 38) = Mid(rev_lpvaluename, 1, 38) Then
          logo_IE7_install_guid_2 = "Y"
          logo_IE7_install_index_2 = correct_cmd_count
      End If
      If Mid(rev_lpvaluename, 1, 1) = "{" Then
          key_strresult = VB.Left(rev_lpvaluename, 38)
          cmd_ext_array(correct_cmd_count).ea_extension_name = key_strresult
          correct_cmd_count = correct_cmd_count + 1
      End If
  Next
 End If
  IE7_logo_next_avail_id = 8192
  If Not rk Is Nothing Then
      rk.Close()
  End If
  rev_dwindex = 0
  ReDim guid_sort_array(correct_cmd_count - 1)
  For rev_dwindex = 0 To correct_cmd_count - 1
      guid_sort_array(rev_dwindex) = cmd_ext_array(rev_dwindex).ea_extension_name
  Next
  Dim mycomparer As New myReverserClass()
  Array.Sort(guid_sort_array, mycomparer)
  If correct_cmd_count > 0 Then
      For rev_dwindex = 0 To correct_cmd_count - 1
          logo_sub_valuename = guid_sort_array(rev_dwindex)
          cmd_ext_array(rev_dwindex).ea_extension_name = logo_sub_valuename
          logo_return_value = setie7cmdmapnum(logo_sub_valuename, IE7_logo_next_avail_id)
          IE7_logo_next_avail_id = IE7_logo_next_avail_id + 1
      Next
      logo_return_value = setie7cmdmapnum("NextId", IE7_logo_next_avail_id)
  End If
End Sub
Private Sub check_for_ie7_toolbar()
  logo_IE7_lreturnvalue = 0
  logo_IE7_subkey = "Software\Microsoft\Internet Explorer\LowRegistry\CommandBar"
  If IE8_flag = True Or IE9_flag = True Then
      logo_IE7_sub_valuename = "CommandBandLayout"
  Else
      logo_IE7_sub_valuename = "ButtonLayout"
  End If
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, False)
  Catch
    Err.Clear()
    logo_IE7_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_IE7_lreturnvalue = 2
  Else
      logo_IE7_lreturnvalue = 0
  End If
  If logo_IE7_lreturnvalue = ERROR_FILE_NOT_FOUND Then
      netnewpath = "Software\Microsoft\Internet Explorer\LowRegistry\CommandBar"
      If Not create_rk Is Nothing Then
        create_rk.Close()
      End If
      Try
        create_rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(netnewpath)
      Catch
        Err.Clear()
        create_prob_flag = True
        logo_IE7_lreturnvalue = 2
      End Try
      Try
        rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, False)
      Catch
        Err.Clear()
        create_prob_flag = True
      End Try
      If rk Is Nothing Then
          logo_IE7_lreturnvalue = 2
      Else
          logo_IE7_lreturnvalue = 0
      End If
  End If
  If logo_IE7_lreturnvalue = 0 Then
   ie7_value_returned = rk.GetValue(logo_IE7_sub_valuename, "notfound")
   If ie7_value_returned.ToString = "notfound" Then
       logo_IE7_lreturnvalue = 2
   Else
       If ie7_value_returned.GetType.ToString = "System.Byte[]" Then
        logo_IE7_lreturnvalue = 0
        logo_IE7_byte_string = ie7_value_returned
        logo_IE7_string_ret_length = logo_IE7_byte_string.GetLength(0)
        If logo_install_activity = True Then
            If logo_IE7_string_ret_length > 4 Then
                logo_IE7_toolbar_remain = (logo_IE7_string_ret_length Mod 4)
                logo_IE7_string_ret_length = logo_IE7_string_ret_length - 1
            End If
        End If
       Else
         logo_IE7_lreturnvalue = 2
       End If
   End If
   If logo_IE7_lreturnvalue = ERROR_FILE_NOT_FOUND Then
       logo_IE7_toolbar_cmdmap_exists = "N"
   Else
       If logo_IE7_lreturnvalue = 0 Or logo_IE7_lreturnvalue = ERROR_MORE_DATA Then
           If logo_IE7_string_ret_length > 4 Then
              logo_IE7_toolbar_cmdmap_exists = "Y"
              ReDim logo_IE7_default_toolbar_byte(logo_IE7_string_ret_length + 84)
              logo_IE7_default_toolbar_byte = logo_IE7_byte_string
           Else
               logo_IE7_toolbar_cmdmap_exists = "N"
           End If
       End If
   End If
  End If
End Sub
Private Sub match_b4_IE7_extensions()
  If IE9_flag = True Then
    cmd_id_add_value = cmd_id_increment_value_IE9
  Else
    If IE8_flag = True Then
        cmd_id_add_value = cmd_id_increment_value_IE8
    Else
        cmd_id_add_value = cmd_id_increment_value_IE7
    End If
  End If
  load_extension_list()
  ReDim guid_sort_array(cmdcount_2 - 1)
  For rev_dwindex = 0 To cmdcount_2 - 1
      guid_sort_array(rev_dwindex) = cmd_IE7_b4_array(rev_dwindex).ce7_valuename
  Next
  Array.Copy(cmd_IE7_b4_array, cmd_IE7_array_sort, cmdcount_2)
  Dim mycomparer3 As New myReverserClass()
  Array.Sort(guid_sort_array, mycomparer3)
  ReDim cmd_IE7_b4_array(90)
  For rev_dwindex = 0 To cmdcount_2 - 1
      IE7_search_guid = guid_sort_array(rev_dwindex).ToString
      For cmdmap_loop = 0 To cmdcount_2 - 1
          If IE7_search_guid = cmd_IE7_array_sort(cmdmap_loop).ce7_valuename Then
              cmd_IE7_b4_array(rev_dwindex).ce7_b4_cmdvalue = _
              cmd_IE7_array_sort(cmdmap_loop).ce7_b4_cmdvalue
              cmd_IE7_b4_array(rev_dwindex).ce7_b4_inst_pos = _
              cmd_IE7_array_sort(cmdmap_loop).ce7_b4_inst_pos
              cmd_IE7_b4_array(rev_dwindex).ce7_b4_mapped_flag = _
              cmd_IE7_array_sort(cmdmap_loop).ce7_b4_mapped_flag
              cmd_IE7_b4_array(rev_dwindex).ce7_enabled_flag = _
              cmd_IE7_array_sort(cmdmap_loop).ce7_enabled_flag
              cmd_IE7_b4_array(rev_dwindex).ce7_IE7_cmdvalue = _
              cmd_IE7_array_sort(cmdmap_loop).ce7_IE7_cmdvalue
              cmd_IE7_b4_array(rev_dwindex).ce7_special_ext = _
              cmd_IE7_array_sort(cmdmap_loop).ce7_special_ext
              cmd_IE7_b4_array(rev_dwindex).ce7_valuelength = _
              cmd_IE7_array_sort(cmdmap_loop).ce7_valuelength
              cmd_IE7_b4_array(rev_dwindex).ce7_valuename = _
              cmd_IE7_array_sort(cmdmap_loop).ce7_valuename
              Exit For
          End If
      Next
  Next
  'If cmdcount_2 > 1 Then'
   If cmdcount_2 >= 1 Then              '11/26/2013 - bug fix for when only 1 cmdmapping exists'
      For rev_dwindex = 0 To cmdcount_2 - 1
          rev_lpvaluename = Space(200)
          rev_lpcbvaluename = 38
          rev_lpvaluename = cmd_IE7_b4_array(rev_dwindex).ce7_valuename
          rev_lpcbvaluename = cmd_IE7_b4_array(rev_dwindex).ce7_valuelength
          IE7_search_length = rev_lpcbvaluename
          If Mid(rev_lpvaluename, 1, 1) = "{" Then
              logo_sub_valuename = Mid(rev_lpvaluename, 1, rev_lpcbvaluename)
              IE7_search_guid = logo_sub_valuename
              If IE8_flag = True Or IE9_flag = True Then
                  search_IE8_extensions()
              Else
                  search_HKCU_extensions()
                  If logo_IE7_found_ext_key = False Then
                      logo_sub_valuename = Mid(rev_lpvaluename, 1, rev_lpcbvaluename)
                      search_HKLM_extensions()
                  End If
              End If
          End If
          logo_IE7_found_ext_key = False
      Next
  End If
  If IE7_b4_cmdcount_hklm > 0 Then
      For IE7_b4_cmdcount_hklm_loop = IE7_b4_cmdcount To ((IE7_b4_cmdcount + IE7_b4_cmdcount_hklm) - 1)
          cmd_IE7_b4_array_hklm(IE7_b4_cmdcount_hklm_loop).ce7_b4_cmdvalue = (IE7_b4_cmdcount_hklm_both + cmd_id_add_value)
          cmd_IE7_b4_array(IE7_b4_cmdcount_hklm_loop).ce7_valuename = _
          cmd_IE7_b4_array_hklm(IE7_b4_cmdcount_hklm_both).ce7_valuename
          cmd_IE7_b4_array(IE7_b4_cmdcount_hklm_loop).ce7_valuelength = _
          cmd_IE7_b4_array_hklm(IE7_b4_cmdcount_hklm_both).ce7_valuelength
          cmd_IE7_b4_array(IE7_b4_cmdcount_hklm_loop).ce7_b4_cmdvalue = (IE7_b4_cmdcount_hklm_loop + cmd_id_add_value)
          IE7_b4_cmdcount_hklm_both = IE7_b4_cmdcount_hklm_both + 1
      Next
      IE7_b4_cmdcount_hklm_both = IE7_b4_cmdcount_hklm_loop
  Else
      IE7_b4_cmdcount_hklm_both = IE7_b4_cmdcount
  End If
  If logo_install_activity = False Then
      match_b4_cmdmaps()
  End If
End Sub
Private Sub match_IE7_extensions()
  If IE9_flag = True Then
    cmd_id_add_value = cmd_id_increment_value_IE9
  Else
    If IE8_flag = True Then
        cmd_id_add_value = cmd_id_increment_value_IE8
    Else
        cmd_id_add_value = cmd_id_increment_value_IE7
    End If
  End If
  logo_b4_proc_flag = False
  logo_IE7_found_ext_key = False
  IE7_cmdcount = 0
  rebuild_IE7_cmdmapping()
  reload_extension_list()
  ReDim guid_sort_array(cmdcount - 1)
  For rev_dwindex = 0 To cmdcount - 1
      guid_sort_array(rev_dwindex) = cmdarray(rev_dwindex).ce_valuename
  Next
  Array.Copy(cmdarray, cmd_IE7_array_sort_2, cmdcount)
  Dim mycomparer4 As New myReverserClass()
  Array.Sort(guid_sort_array, mycomparer4)
  ReDim cmdarray(90)
  For rev_dwindex = 0 To cmdcount - 1
      IE7_search_guid = guid_sort_array(rev_dwindex).ToString
      For cmdmap_loop = 0 To cmdcount - 1
          If IE7_search_guid = cmd_IE7_array_sort_2(cmdmap_loop).ce_valuename Then
              cmdarray(rev_dwindex).ce_b4_cmdvalue = _
              cmd_IE7_array_sort_2(cmdmap_loop).ce_b4_cmdvalue
              cmdarray(rev_dwindex).ce_b4_inst_pos = _
              cmd_IE7_array_sort_2(cmdmap_loop).ce_b4_inst_pos
              cmdarray(rev_dwindex).ce_mapped_flag = _
              cmd_IE7_array_sort_2(cmdmap_loop).ce_mapped_flag
              cmdarray(rev_dwindex).ce_enabled_flag = _
              cmd_IE7_array_sort_2(cmdmap_loop).ce_enabled_flag
              cmdarray(rev_dwindex).ce_IE7_cmdvalue = _
              cmd_IE7_array_sort_2(cmdmap_loop).ce_IE7_cmdvalue
              cmdarray(rev_dwindex).ce_special_ext = _
              cmd_IE7_array_sort_2(cmdmap_loop).ce_special_ext
              cmdarray(rev_dwindex).ce_valuelength = _
              cmd_IE7_array_sort_2(cmdmap_loop).ce_valuelength
              cmdarray(rev_dwindex).ce_valuename = _
              cmd_IE7_array_sort_2(cmdmap_loop).ce_valuename
              cmdarray(rev_dwindex).ce_valuedword = _
              cmd_IE7_array_sort_2(cmdmap_loop).ce_valuedword
              cmdarray(rev_dwindex).ce_valuehex = _
              cmd_IE7_array_sort_2(cmdmap_loop).ce_valuehex
              Exit For
          End If
      Next
  Next
  If cmdcount > 1 Then
      For rev_dwindex = 0 To cmdcount - 1
          rev_lpvaluename = Space(200)
          rev_lpcbvaluename = 38
          rev_lpvaluename = cmdarray(rev_dwindex).ce_valuename
          rev_lpcbvaluename = cmdarray(rev_dwindex).ce_valuelength
          If Mid(rev_lpvaluename, 1, 1) = "{" Then
              logo_sub_valuename = Mid(rev_lpvaluename, 1, rev_lpcbvaluename)
              IE7_search_guid = logo_sub_valuename
              IE7_search_length = rev_lpcbvaluename
              If IE8_flag = True Or IE9_flag = True Then
                  search_IE8_extensions()
              Else
                  search_HKCU_extensions()
                  If logo_IE7_found_ext_key = False Then
                      search_HKLM_extensions()
                  End If
              End If
          End If
          logo_IE7_found_ext_key = False
      Next
  End If
  If IE7_cmdcount_hklm > 0 Then
      For IE7_cmdcount_hklm_loop = IE7_cmdcount To ((IE7_cmdcount + IE7_cmdcount_hklm) - 1)
          cmd_IE7_array(IE7_cmdcount_hklm_loop).ce_valuename = _
           cmd_IE7_b4_array_hklm(IE7_cmdcount_hklm_both).ce7_valuename
          cmd_IE7_array(IE7_cmdcount_hklm_loop).ce_valuelength = _
          cmd_IE7_b4_array_hklm(IE7_cmdcount_hklm_both).ce7_valuelength
          cmd_IE7_array(IE7_cmdcount_hklm_loop).ce_IE7_cmdvalue = (IE7_cmdcount_hklm_loop + cmd_id_add_value)
          IE7_cmdcount_hklm_both = IE7_cmdcount_hklm_both + 1
      Next
      IE7_cmdcount_hklm_both = IE7_cmdcount_hklm_loop
  Else
      IE7_cmdcount_hklm_both = IE7_cmdcount
  End If
End Sub
Private Sub load_extension_list()
  logo_IE7_lreturnvalue = 0
  logo_sub_keyname = "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
  logo_sub_valuename = ""
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
  Catch
    Err.Clear()
    logo_IE7_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_IE7_lreturnvalue = 2
  Else
      logo_IE7_lreturnvalue = 0
  End If
  If logo_IE7_lreturnvalue = 0 Then
   logo_rqve_lpcvalues = rk.ValueCount
   rev_dwindex = 0
   rev_lpvaluename = Space(200)
   rev_lpcbvaluename = 200
   cmdcount_2 = 0
   IE7_b4_cmdcount = 0
   IE7_b4_cmdcount_hklm = 0
   If logo_rqve_lpcvalues > 1 Then
      Try
        rk.GetValueNames()
      Catch
        Err.Clear()
        create_prob_flag = True
      End Try
      For Each Lowvaluename As String In rk.GetValueNames
          rev_lpvaluename = Lowvaluename.ToString
          If Mid(rev_lpvaluename, 1, 1) = "{" Then
              cmd_IE7_b4_array(cmdcount_2).ce7_valuename = rev_lpvaluename
              rev_lpcbvaluename = Lowvaluename.Length
              cmd_IE7_b4_array(cmdcount_2).ce7_valuelength = rev_lpcbvaluename
              cmdcount_2 = cmdcount_2 + 1
          End If
       Next
       logo_b4_cmdmapping_flag = True
   Else
       logo_b4_cmdmapping_flag = False
   End If
  End If
  If Not rk Is Nothing Then
      rk.Close()
  End If
End Sub
Private Sub rebuild_IE7_cmdmapping()
  Delete_Logo_Key("Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping", HKEY_CURRENT_USER)
  check_for_IE7_cmdmapping()
  build_IE7_cmdmapping()
End Sub
Private Sub reload_extension_list()
  logo_IE7_lreturnvalue = 0
  logo_filename_counter = 0
  logo_file_character = ""
  logo_string_returned = ""
  logo_string_ret_length = 200
  logo_sub_keyname = "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping"
  logo_sub_valuename = ""
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
  Catch
    Err.Clear()
    logo_IE7_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_IE7_lreturnvalue = 2
  Else
      logo_IE7_lreturnvalue = 0
  End If
  rev_dwindex = 0
  rev_lpvaluename = Space(200)
  rev_lpcbvaluename = 38
  cmdcount = 0
  IE7_cmdcount = 0
  IE7_cmdcount_hklm = 0
  If logo_IE7_lreturnvalue = 0 Then
     logo_rqve_lpcvalues = rk.ValueCount
  Else
     logo_rqve_lpcvalues = 0
  End If
   If logo_rqve_lpcvalues > 1 Then
      Try
        rk.GetValueNames()
      Catch
        Err.Clear()
        create_prob_flag = True
      End Try
      For Each valueName As String In rk.GetValueNames()
          rev_lpvaluename = Space(200)
          rev_lpcbvaluename = 38
          rev_lpvaluename = valueName.ToString
          If Mid(logo_new_guid, 1, 38) = Mid(rev_lpvaluename, 1, 38) Then
              logo_install_guid = "Y"
              logo_install_index = rev_dwindex
          End If
          If Mid(logo_new_guid_2, 1, 38) = Mid(rev_lpvaluename, 1, 38) Then
              logo_install_guid_2 = "Y"
              logo_install_index_2 = rev_dwindex
          End If
          If Mid(rev_lpvaluename, 1, 1) = "{" Then
              cmdarray(cmdcount).ce_valuename = rev_lpvaluename
              cmdarray(cmdcount).ce_valuelength = rev_lpcbvaluename
              cmdcount = cmdcount + 1
          End If
      Next
      logo_cmdmapping_flag = True
  Else
      logo_cmdmapping_flag = False
  End If
  If Not rk Is Nothing Then
      rk.Close()
  End If
End Sub
Private Sub search_HKLM_extensions()
  logo_IE7_lreturnvalue = 0
  If IE9_flag = True Then                               'IE9 upgrade'
    cmd_id_add_value = cmd_id_increment_value_IE9
  Else
    If IE8_flag = True Then
        cmd_id_add_value = cmd_id_increment_value_IE8
    Else
        cmd_id_add_value = cmd_id_increment_value_IE7
    End If
  End If
  logo_IE7_subkey = "Software\Microsoft\Internet Explorer\Extensions"
  logo_IE7_subkeyname = logo_IE7_subkey & "\" & Mid(IE7_search_guid, 1, 38)
  logo_IE7_subkeyname = Mid(logo_IE7_subkeyname, 1, 86)
  Try
    rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_IE7_subkeyname, False)
  Catch
    Err.Clear()
    logo_IE7_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_IE7_lreturnvalue = 2
  Else
      logo_IE7_lreturnvalue = 0
  End If
  If logo_IE7_lreturnvalue = 0 Then
   logo_string_returned = ""
   logo_sub_valuename = "HotIcon"
   Try
     logo_string_returned = rk.GetValue(logo_sub_valuename, "notfound").ToString
   Catch
     Err.Clear()
     logo_IE7_lreturnvalue = 2
   End Try
   If logo_string_returned = "notfound" Then
       logo_IE7_lreturnvalue = 2
   Else
       logo_IE7_lreturnvalue = 0
   End If
   If logo_IE7_lreturnvalue = 0 Then
       logo_sub_valuename = "Icon"
       logo_string_returned = rk.GetValue(logo_sub_valuename, "notfound").ToString
       If logo_string_returned = "notfound" Then
           logo_IE7_lreturnvalue = 2
       Else
         If ie7_value_returned.GetType.ToString = "System.String" Or _
            ie7_value_returned.GetType.ToString = "System.Int32" Then
              logo_IE7_lreturnvalue = 0
         End If
       End If
       If logo_IE7_lreturnvalue = 0 Then
           If Not rk Is Nothing Then
               rk.Close()
           End If
           check_ext_validity()
           If logo_current_ext_status = "E" Then
             logo_found_in_hive = "HKLM"
             If logo_b4_proc_flag = True Then
               cmd_IE7_b4_array_hklm(IE7_b4_cmdcount_hklm).ce7_valuename = IE7_search_guid
               cmd_IE7_b4_array_hklm(IE7_b4_cmdcount_hklm).ce7_valuelength = 38
               IE7_b4_cmdcount_hklm = IE7_b4_cmdcount_hklm + 1
               logo_IE7_found_ext_key = True
            Else
                cmd_IE7_array_hklm(IE7_cmdcount_hklm).ce_valuename = IE7_search_guid
                cmd_IE7_array_hklm(IE7_cmdcount_hklm).ce_valuelength = 38
                IE7_cmdcount_hklm = IE7_cmdcount_hklm + 1
                logo_IE7_found_ext_key = True
            End If
           End If
           logo_IE7_found_ext_key = True
       End If
      End If
 End If
  If Not rk Is Nothing Then
      rk.Close()
  End If
End Sub
Private Sub search_HKCU_extensions()
  logo_IE7_lreturnvalue = 0
  If IE9_flag = True Then
    cmd_id_add_value = cmd_id_increment_value_IE9
  Else
    If IE8_flag = True Then
        cmd_id_add_value = cmd_id_increment_value_IE8
    Else
        cmd_id_add_value = cmd_id_increment_value_IE7
    End If
  End If
  logo_string_returned = ""
  logo_string_ret_length = 200
  logo_IE7_subkey = "Software\Microsoft\Internet Explorer\Extensions"
  logo_IE7_subkeyname = logo_IE7_subkey & "\" & Mid(IE7_search_guid, 1, 38)
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkeyname, False)
  Catch
    Err.Clear()
    logo_IE7_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_IE7_lreturnvalue = 2
  Else
      logo_IE7_lreturnvalue = 0
  End If
  logo_sub_valuename = "HotIcon"
  If logo_IE7_lreturnvalue = 0 Then
    Try
      logo_string_returned = rk.GetValue(logo_sub_valuename, "notfound").ToString
    Catch
      Err.Clear()
    End Try
   If logo_string_returned = "notfound" Then
       logo_IE7_lreturnvalue = 2
   Else
       If ie7_value_returned.GetType.ToString = "System.String" Or _
            ie7_value_returned.GetType.ToString = "System.Int32" Then
              logo_IE7_lreturnvalue = 0
       End If
   End If
   If logo_IE7_lreturnvalue = 0 Then
       logo_sub_valuename = "Icon"
       Try
         logo_string_returned = rk.GetValue(logo_sub_valuename, "notfound").ToString
       Catch
         Err.Clear()
         logo_IE7_lreturnvalue = 2
       End Try
       If logo_string_returned = "notfound" Then
           logo_IE7_lreturnvalue = 2
       Else
         If ie7_value_returned.GetType.ToString = "System.String" Or _
            ie7_value_returned.GetType.ToString = "System.Int32" Then
              logo_IE7_lreturnvalue = 0
         End If
       End If
       If logo_IE7_lreturnvalue = 0 Then
           If Not rk Is Nothing Then
               rk.Close()
           End If
           check_ext_validity()
           If logo_current_ext_status = "E" Then
               logo_found_in_hive = "HKCU"
               If logo_b4_proc_flag = True Then
                       cmd_IE7_b4_array(IE7_b4_cmdcount).ce7_valuename = IE7_search_guid
                       cmd_IE7_b4_array(IE7_b4_cmdcount).ce7_valuelength = 38
                       cmd_IE7_b4_array(IE7_b4_cmdcount).ce7_b4_cmdvalue = _
                       (IE7_b4_cmdcount + cmd_id_add_value)
                       IE7_b4_cmdcount = IE7_b4_cmdcount + 1
               Else
                   cmd_IE7_array(IE7_cmdcount).ce_valuename = IE7_search_guid
                   cmd_IE7_array(IE7_cmdcount).ce_valuelength = 38
                   cmd_IE7_array(IE7_cmdcount).ce_IE7_cmdvalue = _
                   (IE7_cmdcount + cmd_id_add_value)
                   IE7_cmdcount = IE7_cmdcount + 1
               End If
       End If
        logo_IE7_found_ext_key = True
    Else
    End If
   End If
  End If
  If Not rk Is Nothing Then
      rk.Close()
  End If
End Sub
Private Sub search_IE8_extensions()
  logo_lreturnvalue = 0
  If IE9_flag = True Then                 'IE9 upgrade'
    cmd_id_add_value = cmd_id_increment_value_IE9
  Else
    If IE8_flag = True Then
        cmd_id_add_value = cmd_id_increment_value_IE8
    Else
        cmd_id_add_value = cmd_id_increment_value_IE7
    End If
  End If
  logo_IE7_subkey = "Software\Microsoft\Internet Explorer\Extensions"
  logo_IE7_subkeyname = logo_IE7_subkey & "\" & Mid(IE7_search_guid, 1, 38)
  logo_IE7_subkeyname = Mid(logo_IE7_subkeyname, 1, 86)
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkeyname, True)
  Catch
    Err.Clear()
    logo_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_lreturnvalue = 2
  Else
      logo_lreturnvalue = 0
  End If
  If logo_lreturnvalue = 0 Then
      logo_found_in_hive = "HKCU"
      GoTo s_IE8_good
  End If
  Try
    rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_IE7_subkeyname, True)
  Catch
    Err.Clear()
    logo_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_lreturnvalue = 2
  Else
      logo_lreturnvalue = 0
  End If
  If logo_lreturnvalue = 0 Then
      logo_found_in_hive = "HKLM"
      GoTo s_IE8_good
  End If
  Exit Sub
s_IE8_good:
  If Not rk Is Nothing Then
      rk.Close()
  End If
  check_ext_validity()
  If logo_current_ext_status = "E" Then
    If logo_b4_proc_flag = True Then
        If logo_found_in_hive = "HKCU" Then
            cmd_IE7_b4_array(IE7_b4_cmdcount).ce7_valuename = _
            IE7_search_guid
            cmd_IE7_b4_array(IE7_b4_cmdcount).ce7_valuelength = _
            IE7_search_length
            cmd_IE7_b4_array(IE7_b4_cmdcount).ce7_b4_cmdvalue = (IE7_b4_cmdcount + cmd_id_add_value)
            IE7_b4_cmdcount = IE7_b4_cmdcount + 1
            logo_IE7_found_ext_key = True
        Else
            cmd_IE7_b4_array_hklm(IE7_b4_cmdcount_hklm).ce7_valuename = IE7_search_guid
            cmd_IE7_b4_array_hklm(IE7_b4_cmdcount_hklm).ce7_valuelength = IE7_search_length
            IE7_b4_cmdcount_hklm = IE7_b4_cmdcount_hklm + 1
            logo_IE7_found_ext_key = True
        End If
    Else
        If logo_found_in_hive = "HKCU" Then
            cmd_IE7_array(IE7_cmdcount).ce_valuename = _
              IE7_search_guid
            cmd_IE7_array(IE7_cmdcount).ce_valuelength = _
              IE7_search_length
            cmd_IE7_array(IE7_cmdcount).ce_IE7_cmdvalue = (IE7_cmdcount + cmd_id_add_value)
            IE7_cmdcount = IE7_cmdcount + 1
            logo_IE7_found_ext_key = True
        Else
            cmd_IE7_array_hklm(IE7_cmdcount_hklm).ce_valuename = IE7_search_guid
            cmd_IE7_array_hklm(IE7_cmdcount_hklm).ce_valuelength = IE7_search_length
            IE7_cmdcount_hklm = IE7_cmdcount_hklm + 1
            logo_IE7_found_ext_key = True
        End If
    End If
    logo_IE7_found_ext_key = True
  End If
End Sub
Private Sub check_for_cmdid_delete()
  num_of_IE7_logo_buttons = 0
  logo_b4_toolbar_prev_install = False
  If logo_IE7_string_ret_length > 4 Then
    logo_b4_begin_button_num = logo_IE7_default_toolbar_byte(4)
    logo_b4_num_of_toolbar_items = CInt(((logo_IE7_string_ret_length - 4) / 4))
    For logo_d = 0 To logo_b4_num_of_toolbar_items - 1
      logo_b4_last_button_seq = (logo_d * 4) + 4
      logo_b4_last_button_value = logo_IE7_default_toolbar_byte(logo_b4_last_button_seq)
       If IE7_search_cmd_id = logo_b4_last_button_value Then
          logo_b4_toolbar_prev_install = True
          num_of_b4_logo_buttons = num_of_b4_logo_buttons + 1
          If logo_IE7_allow_multi_installs = False Then
              cmd_IE7_b4_array(match_b4_loop).ce7_b4_inst_pos = logo_b4_last_button_seq
              Exit For
          End If
      End If
    Next
  End If
End Sub
Private Sub check_for_cmdid_install()
  num_of_IE7_logo_buttons = 0
  logo_b4_toolbar_prev_install = False
  If logo_b4_string_ret_length > 4 Then
    logo_b4_begin_button_num = logo_b4_IE7_default_toolbar_byte(4)
    logo_b4_num_of_toolbar_items = CInt(((logo_b4_string_ret_length - 4) / 4))
    For logo_d = 0 To logo_b4_num_of_toolbar_items - 1
      logo_b4_last_button_seq = (logo_d * 4) + 4
      logo_b4_last_button_value = logo_b4_IE7_default_toolbar_byte(logo_b4_last_button_seq)
      If IE7_search_cmd_id = CInt(logo_b4_last_button_value) Then
        logo_b4_toolbar_prev_install = True
        num_of_b4_logo_buttons = num_of_b4_logo_buttons + 1
        If logo_IE7_allow_multi_installs = False Then
            cmd_IE7_b4_array(match_b4_loop).ce7_b4_inst_pos = logo_b4_last_button_seq
            Exit For
        End If
      End If
    Next
  End If
End Sub
Private Sub check_for_prev_IE7_install()
  num_of_IE7_logo_buttons = 0
  If IE9_flag = True Then                  'IE9 upgrade'
     max_poss_cmd_id = IE7_cmdcount_hklm_both + cmd_id_compare_value_IE9
  Else
    If IE8_flag = True Then
        max_poss_cmd_id = IE7_cmdcount_hklm_both + cmd_id_compare_value_IE8
    Else
        max_poss_cmd_id = IE7_cmdcount_hklm_both + cmd_id_compare_value_IE7
    End If
  End If
  match_b4_n_after()
  match_IE7_cmdmaps()
  logo_IE7_toolbar_prev_install = False
  If logo_IE7_string_ret_length > 4 Then
    If logo_IE7_cmdid_installed = True Then
      logo_IE7_begin_button_num = logo_IE7_default_toolbar_byte(4)
      logo_IE7_num_of_toolbar_items = CInt((logo_IE7_string_ret_length - 4) / 4)
      For logo_d = 0 To logo_IE7_num_of_toolbar_items - 1
        logo_IE7_last_button_seq = (logo_d * 4) + 4
        logo_IE7_last_button_value = logo_IE7_default_toolbar_byte(logo_IE7_last_button_seq)
        If IE7_current_cmd_id = logo_IE7_last_button_value Then
          logo_IE7_toolbar_prev_install = True
          num_of_IE7_logo_buttons = num_of_IE7_logo_buttons + 1
        End If
      Next
    End If
  End If
End Sub
Private Sub deinstall_IE7_buttons()
  If IE9_flag = True Then              'IE9 upgrade'
    cmd_id_compare_value = cmd_id_compare_value_IE9
  Else
    If IE8_flag = True Then
        cmd_id_compare_value = cmd_id_compare_value_IE8
    Else
        cmd_id_compare_value = cmd_id_compare_value_IE7
    End If
  End If
  If logo_IE7_string_ret_length >= 8 Then
    logo_IE7_begin_button_num = logo_IE7_default_toolbar_byte(4)
    logo_IE7_num_of_toolbar_items = CInt(((logo_IE7_string_ret_length - 4) / 4))
    num_of_IE7_logo_buttons = 0
    For logo_IE7_i = 0 To logo_IE7_num_of_toolbar_items - 1
      logo_IE7_last_button_seq = (logo_IE7_i * 4) + 4
      logo_IE7_last_button_value = logo_IE7_default_toolbar_byte(logo_IE7_last_button_seq)
      If b4_current_cmd_id = logo_IE7_last_button_value Then
        If num_of_IE7_logo_buttons = 0 Then
            logo_IE7_target_button_pos = logo_IE7_i + 1
        End If
        num_of_IE7_logo_buttons = num_of_IE7_logo_buttons + 1
        Else
          If logo_IE7_last_button_value > cmd_id_compare_value Then
              logo_other_ie7_cmds_flag = True
          End If
        End If
      Next
  End If
  If num_of_IE7_logo_buttons = 1 Then
    If logo_IE7_target_button_pos = 1 Then
        delete_first_IE7_button()
        GoTo check_other_cmds
    End If
    If logo_IE7_target_button_pos = logo_IE7_num_of_toolbar_items Then
        delete_last_IE7_button()
        GoTo check_other_cmds
    End If
    If num_of_IE7_logo_buttons > 0 Then
        For logo_IE7_i = 1 To num_of_IE7_logo_buttons
            delete_new_IE7_button()
            find_new_IE7_button()
        Next
        update_new_IE7_button()
    End If
  Else
    If num_of_IE7_logo_buttons > 0 Then
        For logo_IE7_i = 1 To num_of_IE7_logo_buttons
            delete_new_IE7_button()
            find_new_IE7_button()
        Next
        update_new_IE7_button()
    End If
  End If
check_other_cmds:
   ' exit sub'  
End Sub
Private Sub find_new_IE7_button()
  If logo_IE7_string_ret_length > 8 Then
    logo_IE7_begin_button_num = logo_IE7_default_toolbar_byte(4)
    logo_IE7_num_of_toolbar_items = CInt((logo_IE7_string_ret_length - 4) / 4)
    For logo_IE7_d = 0 To logo_IE7_num_of_toolbar_items - 1
      logo_IE7_last_button_seq = (logo_IE7_d * 4) + 4
      logo_IE7_last_button_value = logo_IE7_default_toolbar_byte(logo_IE7_last_button_seq)
      If b4_current_cmd_id = logo_IE7_last_button_value Then
        If num_of_IE7_logo_buttons = 0 Then
            logo_IE7_target_button_pos = logo_IE7_d + 1
        End If
      End If
    Next
  End If
End Sub
Private Sub delete_first_IE7_button()
  logo_IE7_start_pos = 4
  For logo_d = 8 To (((logo_IE7_num_of_toolbar_items * 4) + 4) - 1)
    logo_IE7_default_toolbar_byte(logo_d - 4) = logo_IE7_default_toolbar_byte(logo_d)
  Next
  logo_IE7_num_of_toolbar_items = logo_IE7_num_of_toolbar_items - 1
  logo_IE7_string_ret_length = logo_IE7_string_ret_length - 4
  Array.Resize(logo_IE7_default_toolbar_byte, logo_IE7_string_ret_length)
  update_new_IE7_button()
End Sub
Private Sub delete_last_IE7_button()
  logo_IE7_start_pos = ((logo_IE7_num_of_toolbar_items * 4) + 4) - 4
  logo_IE7_string_ret_length = logo_IE7_start_pos
  Array.Resize(logo_IE7_default_toolbar_byte, logo_IE7_string_ret_length)
  update_new_IE7_button()
End Sub
Private Sub delete_new_IE7_button()
  logo_IE7_num_of_toolbar_items = CInt(((logo_IE7_string_ret_length - 4) / 4))
  If logo_IE7_target_button_pos > 1 Then
    logo_IE7_button_difference = logo_IE7_num_of_toolbar_items - logo_IE7_target_button_pos
    logo_num_bytes_to_move = logo_IE7_button_difference * 4
    logo_IE7_start_pos = ((logo_IE7_target_button_pos - 1) * 4) + 4
    If logo_IE7_button_difference > 0 Then
      For logo_d = (logo_IE7_start_pos + 4) To (((logo_IE7_num_of_toolbar_items * 4) + 4) - 1)
          logo_IE7_default_toolbar_byte(logo_d - 4) = logo_IE7_default_toolbar_byte(logo_d)
      Next
      logo_IE7_string_ret_length = logo_IE7_string_ret_length - 4
      Array.Resize(logo_IE7_default_toolbar_byte, logo_IE7_string_ret_length)
    End If
  Else
    If logo_IE7_target_button_pos = logo_IE7_num_of_toolbar_items Then
        logo_IE7_string_ret_length = logo_IE7_string_ret_length - 4
        Array.Resize(logo_IE7_default_toolbar_byte, logo_IE7_string_ret_length)
        Exit Sub
    End If
    If logo_IE7_target_button_pos = 1 Then
      logo_IE7_start_pos = 4
      For logo_d = 8 To (((logo_IE7_num_of_toolbar_items * 4) + 4) - 1)
          logo_IE7_default_toolbar_byte(logo_d - 4) = logo_IE7_default_toolbar_byte(logo_d)
      Next
      logo_IE7_num_of_toolbar_items = logo_num_of_toolbar_items - 1
      logo_IE7_string_ret_length = logo_IE7_string_ret_length - 4
      Array.Resize(logo_IE7_default_toolbar_byte, logo_IE7_string_ret_length)
      Exit Sub
    End If
  End If
End Sub
Private Sub update_new_IE7_button()
  logo_IE7_lreturnvalue = 0
  logo_IE7_subkey = "Software\Microsoft\Internet Explorer\LowRegistry\CommandBar"
  If IE8_flag = True Or IE9_flag = True Then
     logo_IE7_string_key_value = "CommandBandLayout"
  Else
     logo_IE7_string_key_value = "ButtonLayout"
  End If
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
  Catch
    Err.Clear()
    create_prob_flag = True
    logo_IE7_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
     logo_IE7_lreturnvalue = 2
  Else
     logo_IE7_lreturnvalue = 0
  End If
  If logo_IE7_lreturnvalue = 0 Then
     Try
       rk.SetValue(logo_IE7_string_key_value, logo_IE7_default_toolbar_byte, RegistryValueKind.Binary)
     Catch
       Err.Clear()
       create_prob_flag = True
     End Try
  End If
  If Not rk Is Nothing Then
     rk.Close()
  End If
End Sub
Private Sub check_ext_validity()
  Dim reg_ext_error As Boolean
  Dim logo_ie_ext_flags_char As String
  Dim logo_string_ret_length_2 As Integer
  logo_IE7_lreturnvalue = 0
  logo_current_ext_status = ""
  reg_ext_error = False
  logo_IE7_subkey = "Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & IE7_search_guid
  logo_IE7_sub_valuename = "Flags"
  logo_IE7_string_key_value = "Flags"
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
  Catch
    Err.Clear()
    create_prob_flag = True
    logo_IE7_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_IE7_lreturnvalue = 2
  Else
      logo_IE7_lreturnvalue = 0
  End If
  If logo_IE7_lreturnvalue = ERROR_FILE_NOT_FOUND Then
      Try
        rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_IE7_subkey, True)
      Catch
        Err.Clear()
        create_prob_flag = True
        logo_IE7_lreturnvalue = 2
      End Try
      If rk Is Nothing Then
          logo_IE7_lreturnvalue = 2
      Else
          logo_IE7_lreturnvalue = 0
      End If
      If logo_IE7_lreturnvalue = 0 Then
          GoTo check_ext_setting3
      Else
          logo_current_ext_status = "E"
          If Not rk Is Nothing Then
              rk.Close()
          End If
          Exit Sub
      End If
  Else
      If logo_IE7_lreturnvalue = 0 Or logo_IE7_lreturnvalue = ERROR_MORE_DATA Then
check_ext_setting3:
   If logo_IE7_lreturnvalue > 0 Then
       reg_ext_error = True
   End If
   If logo_IE7_lreturnvalue = 0 Or logo_IE7_lreturnvalue = ERROR_MORE_DATA Then
    logo_string_returned = ""
    logo_string_ret_length = 8
    logo_string_ret_length_2 = 8
    logo_sub_valuename = logo_IE7_sub_valuename
    Try
      ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
    Catch
      Err.Clear()
      logo_IE7_lreturnvalue = 2
    End Try
    If ie7_value_returned.ToString = "notfound" Then
        logo_IE7_lreturnvalue = 2
    Else
        logo_IE7_lreturnvalue = 0
    End If
    logo_string_ret_length = Len(ie7_value_returned.ToString)
    logo_string_returned = ie7_value_returned.ToString
    logo_ie_ext_flags_value = Mid(logo_string_returned, 1, 4)
    If Mid(logo_ie_ext_flags_value, 1, 2) = "16" Then
      logo_current_ext_status = "E"
    Else
      logo_ie_ext_flags_char = Mid(logo_ie_ext_flags_value, 1, 1)
      If logo_ie_ext_flags_char = "2" Or logo_ie_ext_flags_char = "4" Or logo_ie_ext_flags_char = "6" Or logo_ie_ext_flags_char = "8" Then
        logo_current_ext_status = "E"
      End If
      If logo_ie_ext_flags_char = "1" Or logo_ie_ext_flags_char = "3" Or logo_ie_ext_flags_char = "5" Or logo_ie_ext_flags_char = "7" Or logo_ie_ext_flags_char = "9" Then
        logo_current_ext_status = "D"
      End If
    End If
   End If
 End If
  If Not rk Is Nothing Then
      rk.Close()
  End If
 End If
End Sub
Private Sub install_ext_settings()
  logo_IE7_lreturnvalue = 0
  reg_ext_error = False
  logo_IE7_subkey = "Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & logo_new_guid
  logo_IE7_sub_valuename = "Flags"
  logo_IE7_string_key_value = "Flags"
  Try
    rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_IE7_subkey, True)
  Catch
    Err.Clear()
    create_prob_flag = True
    logo_IE7_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_IE7_lreturnvalue = 2
  Else
      logo_IE7_lreturnvalue = 0
  End If
  If logo_IE7_lreturnvalue = ERROR_FILE_NOT_FOUND Then
    netnewpath = "Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & logo_new_guid
    Try
      rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(netnewpath)
    Catch
      Err.Clear()
      create_prob_flag = True
      logo_IE7_lreturnvalue = 2
    End Try
    Try
      rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_IE7_subkey, True)
    Catch
      Err.Clear()
      create_prob_flag = True
      logo_IE7_lreturnvalue = 2
    End Try
    If rk Is Nothing Then
        logo_IE7_lreturnvalue = 2
    Else
        logo_IE7_lreturnvalue = 0
    End If
    If logo_IE7_lreturnvalue > 0 Then
        reg_ext_error = True
    End If
    If logo_IE7_lreturnvalue = 0 Then
    Try
      rk.SetValue(logo_IE7_string_key_value, 2, RegistryValueKind.DWord)
    Catch
      Err.Clear()
      create_prob_flag = True
      logo_IE7_lreturnvalue = 2
    End Try
      SetUpKeyValue(HKEY_LOCAL_MACHINE, logo_IE7_subkey, "Version", "*")
    End If
    If Not rk Is Nothing Then
        rk.Close()
    End If
  Else
    If logo_IE7_lreturnvalue = 0 Then
       Try
         rk.SetValue(logo_IE7_string_key_value, 2, RegistryValueKind.DWord)
       Catch
         Err.Clear()
         create_prob_flag = True
         logo_IE7_lreturnvalue = 2
       End Try
       SetUpKeyValue(HKEY_LOCAL_MACHINE, logo_IE7_subkey, "Version", "*")
       If Not rk Is Nothing Then
           rk.Close()
       End If
    End If
  End If
  logo_IE7_subkey = "Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & logo_new_guid_2
  Try
    rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_IE7_subkey, True)
  Catch
    Err.Clear()
    create_prob_flag = True
    logo_IE7_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_IE7_lreturnvalue = 2
  Else
      logo_IE7_lreturnvalue = 0
  End If
  If logo_IE7_lreturnvalue = ERROR_FILE_NOT_FOUND Then
    netnewpath = "Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & logo_new_guid_2
    create_rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(netnewpath)
    If create_rk Is Nothing Then
      create_prob_flag = True
      logo_IE7_lreturnvalue = 2
    Else
      logo_IE7_lreturnvalue = 0
    End If
    Try
      rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(logo_IE7_subkey, True)
    Catch
      Err.Clear()
      create_prob_flag = True
      logo_IE7_lreturnvalue = 2
    End Try
    If rk Is Nothing Then
      logo_IE7_lreturnvalue = 2
    Else
      logo_IE7_lreturnvalue = 0
    End If
    If logo_IE7_lreturnvalue > 0 Then
        reg_ext_error = True
    End If
    If logo_IE7_lreturnvalue = 0 Then
      Try
        rk.SetValue(logo_IE7_string_key_value, 2, RegistryValueKind.DWord)
      Catch
        Err.Clear()
        create_prob_flag = True
        logo_IE7_lreturnvalue = 2
      End Try
      SetUpKeyValue(HKEY_LOCAL_MACHINE, logo_IE7_subkey, "Version", "*")
    End If
    If Not rk Is Nothing Then
        rk.Close()
    End If
  Else
    If logo_IE7_lreturnvalue = 0 Then
      Try
        rk.SetValue(logo_IE7_string_key_value, 2, RegistryValueKind.DWord)
      Catch
        Err.Clear()
        create_prob_flag = True
        logo_IE7_lreturnvalue = 2
      End Try
      SetUpKeyValue(HKEY_LOCAL_MACHINE, logo_IE7_subkey, "Version", "*")
      If Not rk Is Nothing Then
          rk.Close()
      End If
    End If
  End If
  logo_IE7_subkey = "Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & logo_new_guid
  logo_IE7_sub_valuename = "Flags"
  logo_IE7_string_key_value = "Flags"
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
  Catch
    Err.Clear()
    create_prob_flag = True
    logo_IE7_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_IE7_lreturnvalue = 2
  Else
      logo_IE7_lreturnvalue = 0
  End If
  If logo_IE7_lreturnvalue = ERROR_FILE_NOT_FOUND Then
    netnewpath = "Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & logo_new_guid
    create_rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(netnewpath)
    If create_rk Is Nothing Then
      create_prob_flag = True
      logo_IE7_lreturnvalue = 2
    Else
      logo_IE7_lreturnvalue = 0
    End If
    Try
      rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
    Catch
      Err.Clear()
      create_prob_flag = True
      logo_IE7_lreturnvalue = 2
    End Try
    If rk Is Nothing Then
        logo_IE7_lreturnvalue = 2
    Else
        logo_IE7_lreturnvalue = 0
    End If
    If logo_IE7_lreturnvalue > 0 Then
        reg_ext_error = True
    End If
    If logo_IE7_lreturnvalue = 0 Then
      Try
        rk.SetValue(logo_IE7_string_key_value, 2, RegistryValueKind.DWord)
      Catch
        Err.Clear()
        create_prob_flag = True
      End Try
      SetUpKeyValue(HKEY_CURRENT_USER, logo_IE7_subkey, "Version", "*")
    End If
    If Not rk Is Nothing Then
        rk.Close()
    End If
  Else
    If logo_IE7_lreturnvalue = 0 Then
      Try
        rk.SetValue(logo_IE7_string_key_value, 2, RegistryValueKind.DWord)
      Catch
        Err.Clear()
        create_prob_flag = True
        logo_IE7_lreturnvalue = 2
      End Try
      SetUpKeyValue(HKEY_CURRENT_USER, logo_IE7_subkey, "Version", "*")
      If Not rk Is Nothing Then
          rk.Close()
      End If
    End If
  End If
  logo_IE7_subkey = "Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & logo_new_guid_2
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
  Catch
    Err.Clear()
    create_prob_flag = True
    logo_IE7_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
      logo_IE7_lreturnvalue = 2
  Else
      logo_IE7_lreturnvalue = 0
  End If
  If logo_IE7_lreturnvalue = ERROR_FILE_NOT_FOUND Then
    netnewpath = "Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & logo_new_guid_2
    create_rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(netnewpath)
    If create_rk Is Nothing Then
      create_prob_flag = True
      logo_IE7_lreturnvalue = 2
    Else
      logo_IE7_lreturnvalue = 0
    End If
    Try
      rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_IE7_subkey, True)
    Catch
      Err.Clear()
      create_prob_flag = True
      logo_IE7_lreturnvalue = 2
    End Try
    If rk Is Nothing Then
        logo_IE7_lreturnvalue = 2
    Else
        logo_IE7_lreturnvalue = 0
    End If
    If logo_IE7_lreturnvalue > 0 Then
        reg_ext_error = True
    End If
    If logo_IE7_lreturnvalue = 0 Then
      Try
        rk.SetValue(logo_IE7_string_key_value, 2, RegistryValueKind.DWord)
      Catch
        Err.Clear()
        create_prob_flag = True
        logo_IE7_lreturnvalue = 2
      End Try
      SetUpKeyValue(HKEY_CURRENT_USER, logo_IE7_subkey, "Version", "*")
    End If
    If Not rk Is Nothing Then
        rk.Close()
    End If
  Else
    If logo_IE7_lreturnvalue = 0 Then
      Try
        rk.SetValue(logo_IE7_string_key_value, 2, RegistryValueKind.DWord)
      Catch
        Err.Clear()
        create_prob_flag = True
        logo_IE7_lreturnvalue = 2
      End Try
      SetUpKeyValue(HKEY_CURRENT_USER, logo_IE7_subkey, "Version", "*")
      If Not rk Is Nothing Then
          rk.Close()
      End If
    End If
  End If
End Sub
Private Sub deinstall_reg_entries()
  logo_lreturnvalue = 0
  If IE7_flag = False Then
    logo_guid_map_num_2 = getcmdmapnum(logo_new_guid_2)
    dest_string = Hex(logo_guid_map_num_2)
    dest_hi_string = Mid(dest_string, 3, 2)
    dest_lo_string = Mid(dest_string, 1, 2)
    dest_lo_value = hex2lng(dest_hi_string)
    dest_hi_value = hex2lng(dest_lo_string)
  End If
  deinst_toolbar_flag = True
  logo_b4_proc_flag = True
  logo_nextid_missing_flag = False
  If Me.win2kxp_flag = "N" Then
    logo_sub_keyname = "Software\Microsoft\Internet Explorer\Toolbar"
    logo_sub_valuename = "{1E796980-9CC5-11D1-A83F-00C04FC99D61}"
    logo_sub_valuename = Mid(logo_sub_valuename, 1, 38)
    Try
      rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
    Catch
      Err.Clear()
      logo_lreturnvalue = 2
    End Try
    If rk Is Nothing Then
        logo_lreturnvalue = 2
    Else
        logo_lreturnvalue = 0
    End If
    If logo_lreturnvalue = 0 Then
      Try
        ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
      Catch
        Err.Clear()
      End Try
      If ie7_value_returned.ToString = "notfound" Then
        logo_lreturnvalue = 2
      Else
        logo_lreturnvalue = 0
      End If

      If ie7_value_returned.ToString = "notfound" Then
         logo_toolbar_cmdmap_exists = "N"
      Else
         If logo_lreturnvalue = 0 Or logo_lreturnvalue = ERROR_MORE_DATA Then
            logo_toolbar_cmdmap_exists = "Y"
         End If
      End If
    Else
      logo_toolbar_cmdmap_exists = "N"
    End If
  Else
    logo_sub_keyname = "Software\Microsoft\Internet Explorer\Toolbar"
    logo_sub_valuename = "{1E796980-9CC5-11D1-A83F-00C04FC99D61}"
    logo_sub_valuename = Mid(logo_sub_valuename, 1, 38)
    Try
      rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
    Catch
      Err.Clear()
      logo_lreturnvalue = 2
    End Try

    If rk Is Nothing Then
        logo_lreturnvalue = 2
    Else
        logo_lreturnvalue = 0
    End If
    If logo_lreturnvalue = 0 Then
        ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
        If ie7_value_returned.ToString = "notfound" Then
            logo_lreturnvalue = 2
        Else
            logo_lreturnvalue = 0
        End If
        If logo_lreturnvalue = 2 Then
            logo_toolbar_cmdmap_exists = "N"
        Else
            If logo_lreturnvalue = 0 Or logo_lreturnvalue = ERROR_MORE_DATA Then
                logo_toolbar_cmdmap_exists = "Y"
            End If
        End If
    Else
        If logo_lreturnvalue = 2 Then
            logo_toolbar_cmdmap_exists = "N"
        End If
    End If
    If IE7_flag = True Then
        check_for_IE7_cmdmapping()
        correct_IE7_cmdmapping()
        build_IE7_cmdmapping()
    End If
    If Me.win2kxp_flag = "Y" Then
        If IE7_flag = True Then
            check_for_b4_toolbar()
            match_b4_IE7_extensions()
            If logo_b4_toolbar_cmdmap_exists = "Y" Then
                check_for_ie7_toolbar()
                deinstall_IE7_buttons()
            End If
        End If
    End If
  End If
  If Me.win2kxp_flag = "N" Then
    logo_sub_keyname = "Software\Microsoft\Internet Explorer\Toolbar"
    logo_sub_valuename = "{1E796980-9CC5-11D1-A83F-00C04FC99D61}"
    logo_sub_valuename = Mid(logo_sub_valuename, 1, 38)
    Try
      rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
    Catch
      Err.Clear()
      logo_lreturnvalue = 2
    End Try
    If rk Is Nothing Then
        logo_lreturnvalue = 2
    Else
        logo_lreturnvalue = 0
    End If
    logo_string_returned = ""
    If logo_lreturnvalue = 0 Then
      Try
        ie6_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
      Catch
        Err.Clear()
      End Try
      If ie6_value_returned.ToString = "notfound" Then
        logo_lreturnvalue = 2
      Else
        logo_lreturnvalue = 0
        logo_default_toolbar_byte = ie6_value_returned
        logo_string_ret_length = logo_default_toolbar_byte.GetLength(0)

        logo_string_ret_length = (logo_string_ret_length - 1)
        logo_string_returned = ""
      End If
    End If
    If Not rk Is Nothing Then
        rk.Close()
    End If
    If logo_lreturnvalue = 0 Then
        ReDim logo_default_toolbar_byte(logo_string_ret_length + 84)
        Try
          rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
        Catch
          Err.Clear()
          logo_lreturnvalue = 2
        End Try
        If rk Is Nothing Then
            logo_lreturnvalue = 2
        Else
            logo_lreturnvalue = 0
        End If
        If logo_lreturnvalue = 0 Then
          Try
            ie6_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
          Catch
            Err.Clear()
          End Try
          If ie6_value_returned.ToString = "notfound" Then
            logo_lreturnvalue = 2
          Else
            logo_lreturnvalue = 0
          End If
          logo_default_toolbar_byte = ie6_value_returned
          logo_string_returned = ""
          deinst_toolbar_flag = True
        End If
    Else
        If logo_lreturnvalue = ERROR_FILE_NOT_FOUND Then
            deinst_toolbar_flag = False
        End If
    End If
  Else
    logo_sub_keyname = "Software\Microsoft\Internet Explorer\Toolbar"
    logo_sub_valuename = "{1E796980-9CC5-11D1-A83F-00C04FC99D61}"
    logo_sub_valuename = Mid(logo_sub_valuename, 1, 38)
    Try
      rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, True)
    Catch
      Err.Clear()
      logo_lreturnvalue = 2
    End Try
    If rk Is Nothing Then
        logo_lreturnvalue = 2
    Else
        logo_lreturnvalue = 0
    End If
    logo_string_returned = ""
    logo_string_ret_length = 4000
    If logo_lreturnvalue = 0 Then
      Try
        ie6_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
      Catch
        Err.Clear()
        logo_lreturnvalue = 2
      End Try
     If ie6_value_returned.ToString = "notfound" Then
         logo_lreturnvalue = 2
     Else
         logo_lreturnvalue = 0
         logo_byte_string = ie6_value_returned
         logo_string_ret_length = logo_byte_string.GetLength(0)
         If logo_string_ret_length > 2 Then
             logo_string_ret_length = logo_string_ret_length - 1
         End If
     End If
    End If
    If logo_lreturnvalue = 0 Then
      ReDim logo_default_toolbar_byte(logo_string_ret_length + 84)
      Try
        ie6_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
      Catch
        Err.Clear()
        logo_lreturnvalue = 2
      End Try
      If ie6_value_returned.ToString = "notfound" Then
          logo_lreturnvalue = 2
      Else
          logo_lreturnvalue = 0
          logo_default_toolbar_byte = ie6_value_returned
          logo_string_ret_length = logo_default_toolbar_byte.GetLength(0)
          If logo_string_ret_length > 2 Then
              logo_string_ret_length = (logo_string_ret_length - 1)
              logo_string_returned = ""
          End If
      End If
        deinst_toolbar_flag = True
    Else
        If logo_lreturnvalue = ERROR_FILE_NOT_FOUND Then
            deinst_toolbar_flag = False
        End If
    End If
  End If
  If deinst_toolbar_flag = False Then
      GoTo delete_std_keys
  End If
  num_of_logo_buttons = 0
  If logo_string_ret_length > 30 Then
      logo_begin_button_num = logo_default_toolbar_byte(4)
      logo_num_of_toolbar_items = CInt(((logo_string_ret_length - 4) / 28))
      num_of_logo_buttons = 0
      For logo_i = 0 To logo_num_of_toolbar_items - 1
          logo_last_button_seq = (logo_i * 28) + 4
          logo_last_button_value = logo_default_toolbar_byte(logo_last_button_seq + 20)
          logo_last_button_value_2 = logo_default_toolbar_byte(logo_last_button_seq + 21)
          If logo_last_button_value = dest_lo_value Then
              If logo_last_button_value_2 = dest_hi_value Then
                  If num_of_logo_buttons = 0 Then
                      logo_target_button_pos = logo_i + 1
                  End If

                  num_of_logo_buttons = num_of_logo_buttons + 1
              End If
          End If
      Next
  Else
      build_default_toolbar()
      GoTo delete_std_keys
  End If
  If dest_lo_value = 0 Then
      If dest_hi_value = 0 Then
          GoTo delete_std_keys
      End If
  End If
delete_IE6:
  If num_of_logo_buttons = 1 Then
      If logo_target_button_pos = 1 Then
          delete_first_button()
          GoTo delete_std_keys
      End If
      If logo_target_button_pos = logo_num_of_toolbar_items Then
          delete_last_button()
          GoTo delete_std_keys
      End If
      If num_of_logo_buttons > 0 Then
          For logo_i = 1 To num_of_logo_buttons
              delete_new_button()
              find_new_button()
          Next
          update_new_button()
      End If
  Else
      If num_of_logo_buttons > 0 Then
          For logo_i = 1 To num_of_logo_buttons
              delete_new_button()
              find_new_button()
          Next
          update_new_button()
      End If
  End If
delete_std_keys:
  If logo_install_tools = True Then
    Delete_Logo_Key("Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid, HKEY_LOCAL_MACHINE)
  End If
  Delete_Logo_Key("Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, HKEY_LOCAL_MACHINE)
  If logo_install_tools = True Then
    Delete_Logo_Key("Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid, HKEY_CURRENT_USER)
  End If
  Delete_Logo_Key("Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, HKEY_CURRENT_USER)
  Delete_Logo_Value("Software\Microsoft\Internet Explorer\Extensions\CmdMapping", HKEY_CURRENT_USER, logo_new_guid)
  Delete_Logo_Value("Software\Microsoft\Internet Explorer\Extensions\CmdMapping", HKEY_CURRENT_USER, logo_new_guid_2)
  check_nextid()
  If logo_install_add_remove = True Then
    Delete_Logo_Key("Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name, HKEY_LOCAL_MACHINE)
  End If
  If IE7_flag = True Then
    Delete_Logo_Value("Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping", HKEY_CURRENT_USER, logo_new_guid)
    Delete_Logo_Value("Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping", HKEY_CURRENT_USER, logo_new_guid_2)
    Delete_Logo_Key("Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & logo_new_guid, HKEY_CURRENT_USER)
    Delete_Logo_Key("Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & logo_new_guid, HKEY_LOCAL_MACHINE)
    Delete_Logo_Key("Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & logo_new_guid_2, HKEY_CURRENT_USER)
    Delete_Logo_Key("Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & logo_new_guid_2, HKEY_LOCAL_MACHINE)
    Delete_Logo_Key("Software\Microsoft\Windows\CurrentVersion\Ext\Stats\" & logo_new_guid, HKEY_CURRENT_USER)
    Delete_Logo_Key("Software\Microsoft\Windows\CurrentVersion\Ext\Stats\" & logo_new_guid_2, HKEY_LOCAL_MACHINE)
    If logo_b4_toolbar_cmdmap_exists = "Y" Then
      check_for_ie7_toolbar()
      match_IE7_extensions()
      match_for_delete()
    End If
  Else
  End If
  success_delete_msg()
End Sub
Private Sub delete_first_button()
  logo_start_pos = 4
  For logo_d = 32 To (((logo_num_of_toolbar_items * 28) + 4) - 1)
    logo_default_toolbar_byte(logo_d - 28) = logo_default_toolbar_byte(logo_d)
  Next
  logo_num_of_toolbar_items = logo_num_of_toolbar_items - 1
  logo_string_ret_length = logo_string_ret_length - 28
  update_new_button()
End Sub
Private Sub delete_last_button()
  logo_start_pos = ((logo_num_of_toolbar_items * 28) + 4) - 28
  logo_string_ret_length = logo_start_pos
  If logo_string_ret_length > 28 Then
    logo_string_ret_length = logo_string_ret_length - 1
  End If
  Array.Resize(logo_default_toolbar_byte, logo_string_ret_length)
  update_new_button()
End Sub
Private Sub delete_new_button()
  logo_num_of_toolbar_items = CInt((logo_string_ret_length - 4) / 28)
  If logo_target_button_pos > 1 Then
    logo_button_difference = logo_num_of_toolbar_items - logo_target_button_pos
    logo_num_bytes_to_move = logo_button_difference * 28
    logo_start_pos = ((logo_target_button_pos - 1) * 28) + 4
    If logo_button_difference > 0 Then
      For logo_d = (logo_start_pos + 28) To (((logo_num_of_toolbar_items * 28) + 4) - 1)
        logo_default_toolbar_byte(logo_d - 28) = logo_default_toolbar_byte(logo_d)
      Next
      logo_string_ret_length = logo_string_ret_length - 28
    End If
  Else
    If logo_target_button_pos = logo_num_of_toolbar_items Then
      logo_string_ret_length = logo_string_ret_length - 28
      Exit Sub
    End If
    If logo_target_button_pos = 1 Then
      logo_start_pos = 4
      For logo_d = 32 To ((logo_num_of_toolbar_items * 28) + 4)
        logo_default_toolbar_byte(logo_d - 28) = logo_default_toolbar_byte(logo_d)
      Next
      logo_num_of_toolbar_items = logo_num_of_toolbar_items - 1
      logo_string_ret_length = logo_string_ret_length - 28
      Exit Sub
    End If
  End If
End Sub
Private Sub update_new_button()
  logo_lreturnvalue = 0
  logo_seq_button_value = &HF6S
  logo_seq_button_value_2 = &H3S
  If logo_string_ret_length > 30 Then
    logo_begin_button_num = logo_default_toolbar_byte(4)
    logo_num_of_toolbar_items = CInt(((logo_string_ret_length - 4) / 28))
    For logo_d = 0 To logo_num_of_toolbar_items - 1
      logo_last_button_seq = (logo_d * 28) + 4
      logo_last_button_value = logo_default_toolbar_byte(logo_last_button_seq)
      logo_last_button_value_2 = logo_default_toolbar_byte(logo_last_button_seq + 1)
      If logo_last_button_value = 255 Or _
       logo_last_button_value = 254 Then
        If logo_last_button_value_2 = 255 Or _
         logo_last_button_value_2 = 254 Then
        End If
      End If
      If logo_last_button_value < 256 Then
        If logo_last_button_value_2 < 254 Then
          logo_default_toolbar_byte(logo_last_button_seq) = CByte(logo_seq_button_value)
          logo_default_toolbar_byte(logo_last_button_seq + 1) = CByte(logo_seq_button_value_2)
        End If
      End If
      logo_seq_button_value = logo_seq_button_value + 1
      If logo_seq_button_value > 254 Then
        logo_seq_button_value = 0
        If logo_seq_button_value_2 = 3 Then
          logo_seq_button_value_2 = 4
        End If
      End If
    Next
  End If
  logo_subkey = "Software\Microsoft\Internet Explorer\Toolbar"
  logo_string_key_value = "{1E796980-9CC5-11D1-A83F-00C04FC99D61}"
  logo_string_key_value = Mid(logo_string_key_value, 1, 38)
  Try
    rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_subkey, True)
  Catch
    Err.Clear()
    create_prob_flag = True
    logo_lreturnvalue = 2
  End Try
  If rk Is Nothing Then
    logo_lreturnvalue = 2
  Else
    logo_lreturnvalue = 0
  End If
  ReDim Preserve logo_default_toolbar_byte(logo_string_ret_length)
  If logo_lreturnvalue = 0 Then
   Try
   rk.SetValue("{1E796980-9CC5-11D1-A83F-00C04FC99D61}", logo_default_toolbar_byte, RegistryValueKind.Binary)
   Catch
     Err.Clear()
     create_prob_flag = True
   End Try
  End If
  If Not rk Is Nothing Then
    rk.Close()
  End If
End Sub
Private Sub find_new_button()
  If logo_string_ret_length > 30 Then
    logo_begin_button_num = logo_default_toolbar_byte(4)
    logo_num_of_toolbar_items = CInt(((logo_string_ret_length - 4) / 28))
    For logo_d = 0 To logo_num_of_toolbar_items - 1
      logo_last_button_seq = (logo_d * 28) + 4
      logo_last_button_value = logo_default_toolbar_byte(logo_last_button_seq + 20)
      logo_last_button_value_2 = logo_default_toolbar_byte(logo_last_button_seq + 21)
      If logo_last_button_value = dest_lo_value Then
          If logo_last_button_value_2 = dest_hi_value Then
              logo_target_button_pos = logo_d + 1
          End If
      End If
    Next
  End If
End Sub
Private Sub install_reg_entries()
  Dim lretval As Integer
  lretval = 0
  logo_lreturnvalue = 0
  logo_IE7_lreturnvalue = 0
  logo_create_rc = 0
  logo_nextid_spread_changed = False               '01/27/2011'
  locate_favorites()
  get_programs_location()
  logo_auth_copy_system = False
  logo_auth_copy_programs = False
  logo_source_not_found = False
  logo_nextid_missing_flag = False
  logo_exe_path = Application.ExecutablePath()    '08/30/2011'
  logo_source_file = logo_exe_path                '08/03/2011'
  If VB.Right(logo_system_dir, 1) = "\" Then
      logo_dest_file = logo_system_dir & logo_install_module
  Else
      logo_dest_file = logo_system_dir & "\" & logo_install_module
  End If
  Try
      System.IO.File.Copy(logo_source_file, logo_dest_file, True)
  Catch
      logo_return_code = 5
  End Try
  If logo_return_code <> 0 Then
      If logo_return_code = 53 Then
          logo_source_not_found = True
      Else
          If logo_return_code = 70 Then
              logo_auth_copy_system = True
          Else
              logo_other_error_flag = True
          End If
      End If
  Else
      logo_source_not_found = False
      logo_app_path = logo_system_dir
      GoTo install_fav_entry
  End If
  Err.Clear()
  If VB.Right(logo_programs_dir, 1) = "\" Then
      logo_dest_file = logo_programs_dir & logo_install_module
  Else
      logo_dest_file = logo_programs_dir & "\" & logo_install_module
  End If
  Err.Clear()
  Try
      System.IO.File.Copy(logo_source_file, logo_dest_file, True)
  Catch
      logo_return_code = 5
  End Try
  If logo_return_code <> 0 Then
      If logo_return_code = 53 Then
          logo_source_not_found = True
      Else
          If logo_return_code = 70 Then
              logo_auth_copy_program = True
          End If
      End If
  Else
      logo_source_not_found = False
      logo_app_path = logo_programs_dir
      GoTo install_fav_entry
  End If
  logo_return_code = Err.Number
  If logo_return_code <> 0 Then
      If logo_return_code = 53 Then
          logo_source_not_found = True
      Else
          If logo_return_code = 70 Then
              logo_auth_copy_program = True
          End If
      End If
  Else
      logo_source_not_found = False
      logo_app_path = logo_programs_dir
      GoTo install_fav_entry
  End If
  Err.Clear()
  If logo_source_not_found = True Then
    If logo_run_from_tempdir = True Then
      x = MsgBox("Is Program Running within Internet Explorer?" & " Please Save to your Computer and Run outside of Internet Explorer." & " This program will now end", MsgBoxStyle.Information, "Please SAVE FirstButton Installer")
      cancel_button_Click()
      Exit Sub
    End If
  Else
    If logo_auth_copy_program = True Then
      x = MsgBox("Is this User authorized to installed to either the Windows Directory " & "or the Program Files Directory?  Please verify authorization and rerun!" & " This program will now end", MsgBoxStyle.Information, "Please verify Authorization")
      cancel_button_Click()
      Exit Sub
    End If
  End If
  Me.logo_error_code = logo_return_code
  logo_error_handler()
  x = MsgBox("Other error occurred while installing." & " Error Code:" & logo_return_code, MsgBoxStyle.Information, "Problem Installing Button")
  cancel_button_Click()
  Exit Sub
install_fav_entry:
  If VB.Right(logo_app_path, 1) = "\" Then
      logo_app_path = logo_app_path
  Else
      logo_app_path = logo_app_path & "\"
  End If
  create_favorites()
  If IE7_flag = True Then
    check_for_IE7_cmdmapping()
    correct_IE7_cmdmapping()
    build_IE7_cmdmapping()
  End If
  If Me.win2kxp_flag = "Y" Then
      If IE7_flag = True Then
          check_for_b4_toolbar()
          match_b4_IE7_extensions()
      End If
  End If
  check_toolbar_unlocked()                   '07/09/2012'
  If IE8_flag = True Or IE9_flag = True Then
    check_cmdbar_enabled()
  End If
  If IE9_flag = True Or IE10_flag = True Then
    check_cmdbar_MINIE_enabled()
  End If
  netnewpath = "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid
  Try
    create_rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(netnewpath)
  Catch
    Err.Clear()
    logo_create_rc = 2
  End Try
  If create_rk Is Nothing Then
    logo_create_rc = 2
    create_prob_flag = True
  Else
    logo_create_rc = 0
  End If
  netnewpath = "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2
  Try
    create_rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(netnewpath)
  Catch
    Err.Clear()
    logo_create_rc = 2
  End Try
  If create_rk Is Nothing Then
    logo_create_rc = 2
    create_prob_flag = True
  Else
    logo_create_rc = 0
  End If
  If logo_install_tools = True Then
    SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid, "CLSID", "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}")
    SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid, "Default Visible", "Yes")
    SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid, "Exec", logo_company_url)
    SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid, "MenuText", logo_company_name)
  End If
  SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "CLSID", "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}")
  SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "Default Visible", "Yes")
  SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "ButtonText", logo_brand_name)
  If IE7_flag = True Then
    SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "HotIcon", (logo_app_path & logo_install_module & ",1"))

    SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "Icon", (logo_app_path & logo_install_module & ",1"))
  Else
   If Me.win2kxp_flag = "Y" Then
     If Me.win2000_flag = "Y" Then
       SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "HotIcon", (logo_app_path & logo_install_module & ",4"))
       SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "Icon", (logo_app_path & logo_install_module & ",3"))
     Else
       SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "HotIcon", (logo_app_path & logo_install_module & ",2"))
       SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "Icon", (logo_app_path & logo_install_module & ",1"))
     End If
      Else
       SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "HotIcon", (logo_app_path & logo_install_module & ",6"))
       SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "Icon", (logo_app_path & logo_install_module & ",5"))
      End If
  End If
  SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "ToolTip", logo_company_url)
  SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "Exec", "http://" & logo_company_url)
  netnewpath = "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid
  Try
    create_rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(netnewpath)
  Catch
    Err.Clear()
    logo_create_rc = 2
  End Try
  If create_rk Is Nothing Then
    logo_create_rc = 2
    create_prob_flag = True
  Else
    logo_create_rc = 0
  End If
  netnewpath = "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2
  Try
    create_rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(netnewpath)
  Catch
    Err.Clear()
    logo_create_rc = 2
  End Try
  If create_rk Is Nothing Then
    logo_create_rc = 2
    create_prob_flag = True
  Else
    logo_create_rc = 0
  End If
  If logo_install_tools = True Then
    SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid, "CLSID", "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}")
    SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid, "Default Visible", "Yes")
    SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid, "Exec", logo_company_url)
    SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid, "MenuText", logo_company_name)
  End If
  SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "CLSID", "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}")
  SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "Default Visible", "Yes")
  SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "ButtonText", logo_brand_name)
  If IE7_flag = True Then
    SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "Icon", (logo_app_path & logo_install_module & ",1"))
    SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "HotIcon", (logo_app_path & logo_install_module & ",1"))
  Else
    If Me.win2kxp_flag = "Y" Then
      If Me.win2000_flag = "Y" Then
        SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "HotIcon", (logo_app_path & logo_install_module & ",4"))
        SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "Icon", (logo_app_path & logo_install_module & ",3"))
      Else
        SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "HotIcon", (logo_app_path & logo_install_module & ",2"))
        SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "Icon", (logo_app_path & logo_install_module & ",1"))
      End If
    Else
      SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "HotIcon", (logo_app_path & logo_install_module & ",6"))
      SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "Icon", (logo_app_path & logo_install_module & ",5"))
    End If
  End If
  SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "ToolTip", logo_company_url)
  SetUpKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & logo_new_guid_2, "Exec", "http://" & logo_company_url)
  If logo_install_add_remove = True Then
    If Me.win2kxp_flag = "N" Then
      netnewpath = "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name
      Try
        create_rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(netnewpath)
      Catch
        Err.Clear()
        logo_create_rc = 2
      End Try
      If create_rk Is Nothing Then
        logo_create_rc = 2
        create_prob_flag = True
      Else
        logo_create_rc = 0
      End If
      SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name, "UninstallString", (logo_app_path & logo_install_module))
      SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name, "DisplayName", logo_company_name)
    Else
      netnewpath = "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name
      Try
        create_rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(netnewpath)
      Catch
        Err.Clear()
        logo_create_rc = 2
      End Try
      If create_rk Is Nothing Then
        logo_create_rc = 2
        create_prob_flag = True
      Else
        logo_create_rc = 0
      End If
      SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name, "DisplayName", logo_company_name)
      SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name, "DisplayIcon", (logo_app_path & logo_install_module))
      SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name, "Publisher", "Liquidity Lighthouse, LLC")
      SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name, "RegCompany", logo_company_name)
      SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name, "HelpLink", ("http://www.liquiditylighthouse.us"))
      SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name, "URLInfoAbout", ("http://" & logo_company_url))
      SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name, "UninstallString", (logo_app_path & logo_install_module))
      SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name, "ModifyPath", (logo_app_path & logo_install_module))
      SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name, "InstallSource", (logo_app_path & logo_install_module))
      SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name, "DisplayVersion", "1.56.0")       '11/10/2013'
    End If
    SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name, "VersionMajor", "1")
    SetUpKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & logo_brand_name, "VersionMinor", "56")      '11/10/2013'
  End If
  logo_filename_counter = 0
  logo_file_character = ""
  logo_string_returned = ""
  logo_string_ret_length = 200
  IE7_cmdcount = 0
'------ rewriting the following section 01/24/2011 to fix the cmdmapping overwrite problem for IE6 '
  If IE7_flag = False Then
    logo_sub_keyname = "Software\Microsoft\Internet Explorer\Extensions\CmdMapping"
    logo_sub_valuename = ""
    Try
      rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
    Catch
      Err.Clear()
      create_prob_flag = True
      logo_lreturnvalue = 2
    End Try
    If rk Is Nothing Then
        logo_lreturnvalue = 2
    Else
        logo_lreturnvalue = 0
    End If
    If logo_lreturnvalue = 2 Or logo_lreturnvalue = 6 Then
       netnewpath = "Software\Microsoft\Internet Explorer\Extensions\CmdMapping"
       Try
         rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(netnewpath)
       Catch
         Err.Clear()
         create_prob_flag = True
       End Try
       If rk Is Nothing Then
         logo_lreturnvalue = 2
       Else
         logo_lreturnvalue = 0
         logo_next_avail_id = 8192
         logo_return_value = setcmdmapnum("NextId", logo_next_avail_id)
         logo_nextid_missing_flag = False
       End If
       If Not rk Is Nothing Then
           rk.Close()
       End If
       logo_sub_keyname = "Software\Microsoft\Internet Explorer\Extensions\CmdMapping"
       logo_sub_valuename = ""
       Try
         rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
       Catch
         Err.Clear()
         create_prob_flag = True
       End Try
       If rk Is Nothing Then
           logo_lreturnvalue = 2
       Else
           logo_lreturnvalue = 0
       End If
    Else  'if the cmdmapping key does exist - 01/24/2011'
        If logo_lreturnvalue = 0 Then
          check_nextid()
          logo_sub_keyname = "Software\Microsoft\Internet Explorer\Extensions\CmdMapping"
          logo_sub_valuename = ""
          Try
            rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
          Catch
            Err.Clear()
            create_prob_flag = True
          End Try
          If rk Is Nothing Then
              logo_lreturnvalue = 2
          Else
              logo_lreturnvalue = 0
          End If
        End If
      End If
      rev_dwindex = 0
      rev_hkey = hKey
      rev_lpvaluename = Space(200)
      rev_lpcbvaluename = 200
      cmdcount = 0
      If logo_lreturnvalue = 0 Then
       logo_rqve_lpcvalues = rk.ValueCount
       If logo_rqve_lpcvalues > 1 Then
        Try
          rk.GetValueNames()
        Catch
          Err.Clear()
          create_prob_flag = True
        End Try
        For Each valueName As String In rk.GetValueNames()
          rev_lpvaluename = Space(200)
          rev_lpcbvaluename = 200
          If valueName.GetType.ToString = "System.String" Then   '01/24/2011'
              rev_lpvaluename = valueName.ToString               '01/24/2011'
          End If                                                 '01/24/2011'
          If logo_rqve_lpcvalues > 89 And cmdcount > 89 Then     '01/24/2011'
            Exit For                                             '01/24/2011'
          End If                                                 '01/24/2011'
          If Mid(logo_new_guid, 1, 38) = Mid(rev_lpvaluename, 1, 38) Then
              logo_install_guid = "Y"
              logo_install_index = rev_dwindex
          End If
          If Mid(logo_new_guid_2, 1, 38) = Mid(rev_lpvaluename, 1, 38) Then
              logo_install_guid_2 = "Y"
              logo_install_index_2 = rev_dwindex
          End If
          If Mid(rev_lpvaluename, 1, 1) = "{" Then
            cmdarray(cmdcount).ce_valuename = rev_lpvaluename
            cmdarray(cmdcount).ce_valuelength = rev_lpcbvaluename
            cmdcount = cmdcount + 1
          Else
            cmdarray(rev_dwindex).ce_valuelength = 0
          End If
        Next
            logo_cmdmapping_flag = True
        Else
            logo_cmdmapping_flag = False
        End If
        If Not rk Is Nothing Then
            rk.Close()
        End If
      End If
      logo_sub_keyname = "Software\Microsoft\Internet Explorer\Extensions\CmdMapping"
      logo_sub_valuename = ""
      Try
        rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
      Catch
        Err.Clear()
        create_prob_flag = True
      End Try
      If rk Is Nothing Then
          logo_lreturnvalue = 2
      Else
          logo_lreturnvalue = 0
      End If
     If cmdcount > 0 Then
       If cmdcount < 91 Then               '01/24/2011'
        For rev_dwindex = 0 To cmdcount - 1
          rev_lpvaluename = Space(200)
          rev_lpcbvaluename = 200
          logo_string_ret_length = 200
          logo_long_returned = 0
          rev_lpvaluename = cmdarray(rev_dwindex).ce_valuename
          rev_lpcbvaluename = cmdarray(rev_dwindex).ce_valuelength
          If Me.win2kxp_flag = "N" Then
           If Me.winverflag = "winnt4" Then
             logo_sub_valuename = rev_lpvaluename
             Try
               ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
             Catch
               Err.Clear()
             End Try
             If ie7_value_returned.ToString = "notfound" Then
                 logo_lreturnvalue = 2
             Else
                 logo_lreturnvalue = 0
                If ie7_value_returned.GetType.ToString = "System.Int32" Or _
                 ie7_value_returned.GetType.ToString = "System.String" Then
                 logo_string_ret_length = ie7_value_returned.ToString.Length
                 If logo_string_ret_length >= 1 Then
                   If IsNumeric(ie7_value_returned.ToString) Then
                     logo_long_returned = CInt(ie7_value_returned.ToString)
                   End If
                 End If
                End If
             End If
           Else
             logo_sub_valuename = cmdarray(rev_dwindex).ce_valuename
             Try
               ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
             Catch
               Err.Clear()
             End Try
             If ie7_value_returned.ToString = "notfound" Then
                 logo_lreturnvalue = 2
             Else
                logo_lreturnvalue = 0
                If ie7_value_returned.GetType.ToString = "System.Int32" Or _
                 ie7_value_returned.GetType.ToString = "System.String" Then
                 logo_string_ret_length = ie7_value_returned.ToString.Length
                 If logo_string_ret_length >= 1 Then
                   If IsNumeric(ie7_value_returned.ToString) Then
                     logo_long_returned = CInt(ie7_value_returned.ToString)
                   End If
                 End If
                End If
             End If
           End If
      End If
       If Me.win2kxp_flag = "Y" Then
           logo_sub_valuename = rev_lpvaluename
         Try
           ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
         Catch
           Err.Clear()
         End Try
           If ie7_value_returned.ToString = "notfound" Then
             logo_lreturnvalue = 2
           Else
             logo_lreturnvalue = 0
             If ie7_value_returned.GetType.ToString = "System.Int32" Or _
                ie7_value_returned.GetType.ToString = "System.String" Then
                logo_string_ret_length = ie7_value_returned.ToString.Length
                If logo_string_ret_length >= 1 Then
                  If IsNumeric(ie7_value_returned.ToString) Then
                    logo_long_returned = CInt(ie7_value_returned.ToString)
                  End If
                End If
               End If
           End If
       End If
      If logo_lreturnvalue = 0 Then
        If ie7_value_returned.GetType.ToString = "System.Int32" Or _
           ie7_value_returned.GetType.ToString = "System.String" Then
           logo_string_ret_length = ie7_value_returned.ToString.Length
          If logo_string_ret_length >= 1 Then
           If IsNumeric(ie7_value_returned.ToString) Then
           logo_long_returned = CInt(ie7_value_returned.ToString)
           cmdarray(rev_dwindex).ce_valuedword = logo_long_returned
           cmdarray(rev_dwindex).ce_valuehex = Hex(logo_long_returned)
           End If
         End If
         End If
        End If
         Next
        End If            '01/24/2011'
      End If
      If Not rk Is Nothing Then
          rk.Close()
      End If
      ReDim cmdmap_sort_array(cmdcount - 1)
      For rev_dwindex = 0 To cmdcount - 1
        If cmdarray(rev_dwindex).ce_valuedword > 0 Then
          cmdmap_sort_array(rev_dwindex) = cmdarray(rev_dwindex).ce_valuedword.ToString
        End If
      Next
      Dim mycomparer2 As New myReverserClass()
      Array.Sort(cmdmap_sort_array, mycomparer2)
      Array.Copy(cmdarray, cmdarray_sort, cmdcount)
      ReDim cmdarray(90)
      For rev_dwindex = 0 To cmdcount - 1
        current_cmdmap_id = CInt(cmdmap_sort_array(rev_dwindex))
        For cmdmap_loop = 0 To cmdcount - 1
          If current_cmdmap_id = cmdarray_sort(cmdmap_loop).ce_valuedword Then
            cmdarray(rev_dwindex).ce_b4_cmdvalue = _
            cmdarray_sort(cmdmap_loop).ce_b4_cmdvalue
            cmdarray(rev_dwindex).ce_b4_inst_pos = _
            cmdarray_sort(cmdmap_loop).ce_b4_inst_pos
            cmdarray(rev_dwindex).ce_enabled_flag = _
            cmdarray_sort(cmdmap_loop).ce_enabled_flag
            cmdarray(rev_dwindex).ce_IE7_cmdvalue = _
            cmdarray_sort(cmdmap_loop).ce_IE7_cmdvalue
            cmdarray(rev_dwindex).ce_special_ext = _
            cmdarray_sort(cmdmap_loop).ce_special_ext
            cmdarray(rev_dwindex).ce_valuedword = _
            cmdarray_sort(cmdmap_loop).ce_valuedword
            cmdarray(rev_dwindex).ce_valuehex = _
            cmdarray_sort(cmdmap_loop).ce_valuehex
            cmdarray(rev_dwindex).ce_valuelength = _
            cmdarray_sort(cmdmap_loop).ce_valuelength
            cmdarray(rev_dwindex).ce_valuename = _
            cmdarray_sort(cmdmap_loop).ce_valuename
            Exit For
          End If
         Next
      Next
      logo_sub_keyname = "Software\Microsoft\Internet Explorer\Extensions\CmdMapping"
      logo_sub_valuename = "NextId"
      Try
        rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
      Catch
        Err.Clear()
        create_prob_flag = True
      End Try
      If rk Is Nothing Then
          logo_lreturnvalue = 2
      Else
          logo_lreturnvalue = 0
      End If
      logo_string_returned = ""
      logo_string_ret_length = 8
      If Me.win2kxp_flag = "Y" Then
        Try
          ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
        Catch
          Err.Clear()
        End Try
        If ie7_value_returned.ToString = "notfound" Then
          logo_lreturnvalue = 2
          logo_nextid_missing_flag = True
        Else
         logo_lreturnvalue = 0
         If ie7_value_returned.GetType.ToString = "System.Int32" Or _
            ie7_value_returned.GetType.ToString = "System.String" Then
          logo_string_ret_length = ie7_value_returned.ToString.Length
          If logo_string_ret_length >= 1 Then
           If IsNumeric(ie7_value_returned.ToString) Then
            logo_long_returned = CInt(ie7_value_returned.ToString)
            logo_next_avail_id = logo_long_returned
           End If
          End If
         End If                                                      '
        End If
        If Not rk Is Nothing Then
          rk.Close()
        End If
      Else
        If Me.winverflag = "winnt4" Then
          Try
            ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
          Catch
            Err.Clear()
          End Try
          If ie7_value_returned.ToString = "notfound" Then
              logo_lreturnvalue = 2
              logo_nextid_missing_flag = True
          Else
              logo_lreturnvalue = 0
              If ie7_value_returned.GetType.ToString = "System.Int32" Or _
                 ie7_value_returned.GetType.ToString = "System.String" Then
              logo_string_ret_length = ie7_value_returned.ToString.Length
                If logo_string_ret_length >= 1 Then
                  If IsNumeric(ie7_value_returned.ToString) Then
                    logo_long_returned = CInt(ie7_value_returned.ToString)
                    logo_next_avail_id = logo_long_returned
                  End If
                End If
               End If
          End If
          If Not rk Is Nothing Then
              rk.Close()
          End If
         Else
           Try
             ie7_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
           Catch
             Err.Clear()
           End Try
           If ie7_value_returned.ToString = "notfound" Then
               logo_lreturnvalue = 2
               logo_nextid_missing_flag = True
           Else
               logo_lreturnvalue = 0
              If ie7_value_returned.GetType.ToString = "System.Int32" Or _
                 ie7_value_returned.GetType.ToString = "System.String" Then
              logo_string_ret_length = ie7_value_returned.ToString.Length
                If logo_string_ret_length >= 1 Then
                  If IsNumeric(ie7_value_returned.ToString) Then
                    logo_long_returned = CInt(ie7_value_returned.ToString)
                    logo_next_avail_id = logo_long_returned
                  End If
                End If
               End If
           End If
           If Not rk Is Nothing Then
               rk.Close()
           End If
          End If
      End If
      If cmdcount > 1 Then
        If logo_next_avail_id = 0 Then
          find_next_avail_id()
        Else
          If logo_next_avail_id < cmdarray(cmdcount - 1).ce_valuedword Then
              logo_next_avail_id = (cmdarray(cmdcount - 1).ce_valuedword) + 1
          End If
          If logo_next_avail_id = cmdarray(cmdcount - 1).ce_valuedword Then
              logo_next_avail_id = cmdarray(cmdcount - 1).ce_valuedword + 1
          End If
        End If
        If logo_next_avail_id < 8192 Then
            find_next_avail_id()
        End If
        If logo_next_avail_id > 8702 Then
            find_next_avail_id()
        End If
      Else
        If cmdcount = 1 Then
            If cmdarray(0).ce_valuedword >= 8192 And _
            cmdarray(0).ce_valuedword <= 8696 Then
              If logo_nextid_spread_changed = False Then     '01/27/2011'
                logo_next_avail_id = cmdarray(0).ce_valuedword + 1
              End If                                         '01/27/2011'
            Else
                logo_next_avail_id = 8192
            End If
        End If
        If logo_next_avail_id < 8192 Then
            logo_next_avail_id = 8192
        End If
        If logo_next_avail_id > 8702 Then
            If cmdcount > 0 Then
                If cmdarray(0).ce_valuedword >= 8192 And _
                cmdarray(0).ce_valuedword <= 8696 Then
                    logo_next_avail_id = cmdarray(0).ce_valuedword + 1
                Else
                    logo_next_avail_id = 8192
                End If
            End If
          End If
      End If
      If cmdcount = 0 Then
        logo_next_avail_id = 8192
      End If
      If logo_install_guid = "Y" Then
          logo_guid_map_num = getcmdmapnum(logo_new_guid)
      End If
      If logo_install_guid_2 = "Y" Then
          logo_guid_map_num_2 = getcmdmapnum(logo_new_guid_2)
      End If
      If logo_install_guid = "N" Then
        If Me.winverflag = "winnt351" Or Me.winverflag = "winnt4" Then
            'continue'
        Else
            lretval = setcmdmapnum(logo_new_guid, logo_next_avail_id)
            logo_guid_map_num = logo_next_avail_id
            logo_next_avail_id = logo_next_avail_id + 1
        End If
      End If
      If logo_install_guid_2 = "N" Then
        lretval = setcmdmapnum(logo_new_guid_2, logo_next_avail_id)
        logo_guid_map_num_2 = logo_next_avail_id
        logo_next_avail_id = logo_next_avail_id + 1
      End If
      If logo_install_guid = "N" Or logo_install_guid_2 = "N" Then
        lretval = setcmdmapnum("NextId", logo_next_avail_id)
        logo_nextid_missing_flag = False
      Else
        If logo_nextid_missing_flag = True Then
          If logo_next_avail_id > 0 Then
            lretval = setcmdmapnum("NextId", logo_next_avail_id)
          End If
        End If
      End If
  End If
  If IE7_flag = True Then
    IE7_logo_next_avail_id = getie7cmdmapnum("NextId")
    If logo_IE7_install_guid = "N" Then
        logo_return_value = setie7cmdmapnum(logo_new_guid, IE7_logo_next_avail_id)
        IE7_logo_next_avail_id = IE7_logo_next_avail_id + 1
    End If
    If logo_IE7_install_guid_2 = "N" Then
        logo_return_value = setie7cmdmapnum(logo_new_guid_2, IE7_logo_next_avail_id)
        IE7_logo_next_avail_id = IE7_logo_next_avail_id + 1
    End If
    If logo_IE7_install_guid = "N" Or logo_IE7_install_guid_2 = "N" Then
        logo_return_value = setie7cmdmapnum("NextId", IE7_logo_next_avail_id)
    End If
  End If
  If Me.win2kxp_flag = "Y" Then
      If IE7_flag = True Then
          match_IE7_extensions()
      End If
  End If
  If IE7_flag = False Then
      If Me.win2kxp_flag = "N" Then
        logo_sub_keyname = "Software\Microsoft\Internet Explorer\Toolbar"
        logo_sub_valuename = "{1E796980-9CC5-11D1-A83F-00C04FC99D61}"
        logo_sub_valuename = Mid(logo_sub_valuename, 1, 38)
        Try
          rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
        Catch
          Err.Clear()
          create_prob_flag = True
        End Try
        If rk Is Nothing Then
            logo_lreturnvalue = 2
        Else
            logo_lreturnvalue = 0
        End If
        If Me.winverflag = "winnt4" Then
          Try
            ie6_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
          Catch
            Err.Clear()
          End Try
          If ie6_value_returned.ToString = "notfound" Then
              logo_lreturnvalue = 2
          Else
            If ie6_value_returned.GetType.ToString = "System.Byte[]" Then
             logo_default_toolbar_byte = ie6_value_returned
             logo_string_ret_length = logo_default_toolbar_byte.GetLength(0)
             If logo_string_ret_length > 2 Then
                 logo_string_ret_length = (logo_string_ret_length - 1)
                 logo_string_returned = ""
             End If
             logo_lreturnvalue = 0
            Else
             logo_lreturnvalue = 2
            End If
          End If
          Else
            Try
              ie6_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
            Catch
              Err.Clear()
            End Try
            If ie6_value_returned.ToString = "notfound" Then
                logo_lreturnvalue = 2
            Else
              If ie6_value_returned.GetType.ToString = "System.Byte[]" Then
               logo_default_toolbar_byte = ie6_value_returned
               logo_string_ret_length = logo_default_toolbar_byte.GetLength(0)
               If logo_string_ret_length > 2 Then
                   logo_string_ret_length = (logo_string_ret_length - 1)
                   logo_string_returned = ""
               End If
               logo_lreturnvalue = 0
              Else
                logo_lreturnvalue = 2
              End If
            End If
          End If
          If logo_lreturnvalue = ERROR_FILE_NOT_FOUND Then
            logo_toolbar_cmdmap_exists = "N"
          Else
            If logo_lreturnvalue = 0 Or logo_lreturnvalue = ERROR_MORE_DATA Then
              If logo_string_ret_length < 31 Then
                  logo_toolbar_cmdmap_exists = "N"
              Else
                  logo_toolbar_cmdmap_exists = "Y"
              End If
            End If
          End If
      Else
          logo_sub_keyname = "Software\Microsoft\Internet Explorer\Toolbar"
          logo_sub_valuename = "{1E796980-9CC5-11D1-A83F-00C04FC99D61}"
          logo_sub_valuename = Mid(logo_sub_valuename, 1, 38)
          logo_string_ret_length = 8
          Try
            rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
          Catch
            Err.Clear()
            create_prob_flag = True
          End Try
          If rk Is Nothing Then
              logo_lreturnvalue = 2
          Else
              logo_lreturnvalue = 0
          End If
          ie6_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
          If ie6_value_returned.ToString = "notfound" Then
            logo_lreturnvalue = 2
          Else
            If ie6_value_returned.GetType.ToString = "System.Byte[]" Then
             logo_default_toolbar_byte = ie6_value_returned
             logo_string_ret_length = logo_default_toolbar_byte.GetLength(0)
             If logo_string_ret_length > 2 Then
               logo_string_ret_length = (logo_string_ret_length - 1)
               logo_string_returned = ""
             End If
             logo_lreturnvalue = 0
            Else
             logo_lreturnvalue = 2
            End If
          End If
          If logo_lreturnvalue = ERROR_FILE_NOT_FOUND Then
              logo_toolbar_cmdmap_exists = "N"
          Else
            If logo_lreturnvalue = 0 Or logo_lreturnvalue = ERROR_MORE_DATA Then
              If logo_string_ret_length < 31 Then
                  logo_toolbar_cmdmap_exists = "N"
              Else
                  logo_toolbar_cmdmap_exists = "Y"
              End If
            End If
          End If
      End If
  End If
  If IE7_flag = True Then
      check_for_ie7_toolbar()
      If logo_IE7_toolbar_cmdmap_exists = "N" Then
          build_IE7_default_toolbar()
      End If
  End If
  If IE7_flag = False Then
      If logo_toolbar_cmdmap_exists = "N" Then
          build_default_toolbar()
      End If
      If Me.win2kxp_flag = "N" Then
        logo_sub_keyname = "Software\Microsoft\Internet Explorer\Toolbar"
        logo_sub_valuename = "{1E796980-9CC5-11D1-A83F-00C04FC99D61}"
        logo_sub_valuename = Mid(logo_sub_valuename, 1, 38)
        Try
          rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
        Catch
          Err.Clear()
          create_prob_flag = True
        End Try
        If rk Is Nothing Then
            logo_lreturnvalue = 2
        Else
            logo_lreturnvalue = 0
        End If
        Try
          ie6_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
        Catch
          Err.Clear()
        End Try
        If ie6_value_returned.ToString = "notfound" Then
            logo_lreturnvalue = 2
        Else
            logo_lreturnvalue = 0
            If ie6_value_returned.GetType.ToString = "System.Byte[]" Then
             logo_default_toolbar_byte = ie6_value_returned
             logo_string_ret_length = logo_default_toolbar_byte.GetLength(0)

             If logo_string_ret_length > 2 Then
                 logo_string_ret_length = (logo_string_ret_length - 1)
                 logo_string_returned = ""
             End If
            End If
        End If
        If logo_lreturnvalue = 0 Then
          If Me.winverflag = "winnt4" Then
            Try
              ie6_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
            Catch
              Err.Clear()
            End Try
            If ie6_value_returned.ToString = "notfound" Then
                logo_lreturnvalue = 2
            Else
                logo_lreturnvalue = 0
            End If
            If logo_lreturnvalue = 0 Then
                logo_default_toolbar_byte = ie6_value_returned
                logo_string_ret_length = logo_default_toolbar_byte.GetLength(0)

                If logo_string_ret_length > 2 Then
                    logo_string_ret_length = (logo_string_ret_length - 1)
                    logo_string_returned = ""
                End If
            End If
            logo_string_returned = ""
          Else
              Try
                ie6_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
              Catch
                Err.Clear()
              End Try
              If ie6_value_returned.ToString = "notfound" Then
                logo_lreturnvalue = 2
              Else
                logo_lreturnvalue = 0
                logo_default_toolbar_byte = ie6_value_returned
                logo_string_ret_length = logo_default_toolbar_byte.GetLength(0)
                If logo_string_ret_length > 2 Then
                    logo_string_ret_length = (logo_string_ret_length - 1)
                    logo_string_returned = ""
                End If
              End If
            End If
        End If
      Else
        logo_sub_keyname = "Software\Microsoft\Internet Explorer\Toolbar"
        logo_sub_valuename = "{1E796980-9CC5-11D1-A83F-00C04FC99D61}"
        logo_sub_valuename = Mid(logo_sub_valuename, 1, 38)
        Try
          rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(logo_sub_keyname, False)
        Catch
          Err.Clear()
          create_prob_flag = True
        End Try
        If rk Is Nothing Then
          logo_lreturnvalue = 2
        Else
          logo_lreturnvalue = 0
        End If
        If logo_lreturnvalue = 0 Then
          Try
            ie6_value_returned = rk.GetValue(logo_sub_valuename, "notfound")
          Catch
            Err.Clear()
          End Try
        End If
        If ie6_value_returned.ToString = "notfound" Then
            logo_lreturnvalue = 2
        Else
            logo_lreturnvalue = 0
        End If
        If logo_lreturnvalue = 0 Then
          logo_default_toolbar_byte = ie6_value_returned
          logo_string_ret_length = logo_default_toolbar_byte.GetLength(0)
          If logo_string_ret_length > 2 Then
            logo_string_ret_length = (logo_string_ret_length - 1)
            logo_string_returned = ""
          End If
        End If
      End If
      ReDim logo_default_toolbar_map(logo_string_ret_length + 84)
      check_for_prev_install()
  End If
  If IE7_flag = True Then
      check_for_prev_IE7_install()
  End If
  If logo_toolbar_prev_install = True Then
      If IE7_flag = False Then
        If logo_allow_multi_installs = False Then
          success_install_msg()
          cancel_button_Click()
          Exit Sub
        End If
      Else
        If logo_IE7_toolbar_prev_install = True Then
          If logo_IE7_allow_multi_installs = False Then
              success_install_msg()
              cancel_button_Click()
              Exit Sub
          End If
        End If
      End If
  Else
      If logo_IE7_toolbar_prev_install = True Then
        If logo_IE7_allow_multi_installs = False Then
          success_install_msg()
          cancel_button_Click()
          Exit Sub
        End If
      End If
  End If
  If IE7_flag = False Then
    If logo_string_ret_length > 30 Then
      logo_begin_button_num = logo_default_toolbar_byte(4)
      logo_num_of_toolbar_items = CInt(((logo_string_ret_length - 4) / 28))
      logo_last_button_value = logo_default_toolbar_byte(logo_last_button_seq)
      logo_last_button_value_2 = logo_default_toolbar_byte(logo_last_button_seq + 1)
      If logo_last_button_value = 255 Or _
        logo_last_button_value = 254 Then
         If logo_last_button_value_2 = 255 Or _
            logo_last_button_value_2 = 254 Then
         End If
      End If
       evaluate_buttons()
      End If
  End If
  If IE7_flag = True Then
    If logo_IE7_string_ret_length > 4 Then
      logo_IE7_num_of_toolbar_items = CInt(((logo_IE7_string_ret_length - 4) / 4))
      evaluate_IE7_buttons()
    End If
  End If
End Sub
Private Sub installation_error()
 x = MsgBox("Installation Error - " & Err.Description & "   " & _
      "  You may not have sufficient authority to run FirstButton on your system.", MsgBoxStyle.Information, _
      "FirstButton  Installation Error")
End Sub
Private Sub tgt_inst_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
  Dim Button As Integer = eventArgs.Button \ &H100000
  Dim Shift As Integer = System.Windows.Forms.Control.ModifierKeys \ &H10000
  Dim x As Double = VB6.PixelsToTwipsX(eventArgs.X)
  Dim y As Double = VB6.PixelsToTwipsY(eventArgs.Y)
End Sub
Private Sub tgt_inst_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
  Dim Cancel As Boolean = eventArgs.Cancel
  Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
  tgt_license.Close()
  eventArgs.Cancel = Cancel
End Sub
Private Sub install_option_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles install_option.CheckedChanged
  If eventSender.Checked Then
    install_option.Checked = True
    uninstall_option.Checked = False
  End If
End Sub
Private Sub next_image_Click()
  If install_option.Checked = True Then
    If agree_check.CheckState > 0 Then
        If logo_check_UAC_flag = True Then    '11/10/2013 - moved here 04/06/2015'
         'check registry to see if UAC is enabled'
         check_UAC_setting()                  '11/10/2013 - moved here 04/06/2015'
        End If                                '11/10/2013 - moved here 04/06/2015'
        If win_UAC_enabled_flag = True Then   '11/10/2013 - moved here 04/06/2015'
         'check to see if this program was launched with Run as Administrator'
         MsgBox("User Account Control is enabled on this computer. If this program was not run with Administor authority, it may fail." + vbCr + vbCr + "If error occurs - please Run As Administrator, outside of Internet Explorer.", _
         MsgBoxStyle.Information, "FirstButton  Installer")          '11/10/2013 - moved here 04/06/2015'
        End If                                '11/10/2013 - moved here 04/06/2015'
      logo_install_activity = True
      install_reg_entries()
    Else
      x = MsgBox("Please check box indicating you accept the User Agreement", MsgBoxStyle.Information, "FirstButton  Installer")
      Exit Sub
    End If
  Else
        If logo_check_UAC_flag = True Then    '11/10/2013 - moved here 04/06/2015'
         'check registry to see if UAC is enabled'
         check_UAC_setting()                  '11/10/2013 - moved here 04/06/2015'
        End If                                '11/10/2013 - moved here 04/06/2015'
        If win_UAC_enabled_flag = True Then   '11/10/2013 - moved here 04/06/2015'
         'check to see if this program was launched with Run as Administrator'
         MsgBox("User Account Control is enabled on this computer. If this program was not run with Administor authority, it may fail." + vbCr + vbCr + "If error occurs - please Run As Administrator, outside of Internet Explorer.", _
         MsgBoxStyle.Information, "FirstButton  Installer")          '11/10/2013 - moved here 04/06/2015'
        End If                                '11/10/2013 - moved here 04/06/2015'
    logo_install_activity = False
    deinstall_reg_entries()
  End If
End Sub
Private Sub next_cmd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles next_cmd.Click
  next_image_Click()
  Exit Sub
End Sub
Private Sub read_agmt_cmd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles read_agmt_cmd.Click
  tgt_license.Show()
End Sub
Private Sub uninstall_option_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles uninstall_option.CheckedChanged
  If eventSender.Checked Then
    install_option.Checked = False
    uninstall_option.Checked = True
  End If
End Sub
End Class