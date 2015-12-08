Option Strict Off
Option Explicit On
Module reghelp
'-------------------------------------------------------------------------------------------'
' THIS PROGRAM IS CONFIDENTIAL AND PROPRIETARY TO LIQUIDITY LIGHTHOUSE, LLC                 '
'  AND MAY NOT BE REPRODUCED, PUBLISHED, OR DISCLOSED TO OTHERS WITHOUT                     '
'   COMPANY AUTHORIZATION.                                                                  '
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
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Integer

Public Structure cmdmapping_entry
   Dim cmd_valuename As String
   Dim cmd_valuedword As Integer
   Dim cmd_index As Integer
End Structure

Public Structure cmdmap_entry
  Dim ce_valuename As String
  Dim ce_valuelength As Integer
  Dim ce_valuedword As Integer
  Dim ce_valuehex As String
  Dim ce_IE7_cmdvalue As Integer
  Dim ce_b4_cmdvalue As Integer
  Dim ce_mapped_flag As Boolean
  Dim ce_b4_inst_pos As Integer
  Dim ce_enabled_flag As Boolean
  Dim ce_special_ext As Boolean
End Structure

 Public Structure ie7_cmdmap_entry
  Dim ce7_valuename As String
  Dim ce7_valuelength As Integer
  Dim ce7_b4_cmdvalue As Integer
  Dim ce7_IE7_cmdvalue As Integer
  Dim ce7_b4_mapped_flag As Boolean
  Dim ce7_b4_inst_pos As Integer
  Dim ce7_enabled_flag As Boolean
  Dim ce7_special_ext As Boolean
 End Structure

 Public Structure extension_array
   Dim ea_extension_name As String
 End Structure

 Public Structure toolbar_map_entry
   Dim tm_value_hi_dword As Integer
   Dim tm_value_lo_dword As Integer
   Dim tm_valuehex As String
 End Structure

Public Structure OSVERSIONINFO
  Dim dwOSVersionInfoSize As Integer
  Dim dwMajorVersion As Integer
  Dim dwMinorVersion As Integer
  Dim dwBuildNumber As Integer
  Dim dwPlatformId As Integer
  <VBFixedString(128), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=128)> Public szCSDVersion() As Char
 End Structure

 Public Const ERROR_ACCESS_DENIED As Short = 5
 Public Const ERROR_FILE_NOT_FOUND As Short = 2
 Public Const ERROR_MORE_DATA As Short = 234
 Public Const ERROR_NO_DATA As Short = 232
 Public Const ERROR_SUCCESS As Short = 0
 Public Const HKEY_CLASSES_ROOT As Integer = &H80000000
 Public Const HKEY_CURRENT_CONFIG As Integer = &H80000005
 Public Const HKEY_CURRENT_USER As Integer = &H80000001
 Public Const HKEY_DYN_DATA As Integer = &H80000006
 Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002

 Public Const cmd_id_cutoff_value_IE7 As Integer = 11
 Public Const cmd_id_cutoff_value_IE8 As Integer = 104
 Public Const cmd_id_cutoff_value_IE9 As Integer = 104
 'Public Const cmd_id_cutoff_value_IE9 As Integer = 117      IE9RC'
 Public Const cmd_id_compare_value_IE7 As Integer = 15
 Public Const cmd_id_compare_value_IE8 As Integer = 104
 Public Const cmd_id_compare_value_IE9 As Integer = 104
 'Public Const cmd_id_compare_value_IE9 As Integer = 117     IE9RC'
 Public Const cmd_id_increment_value_IE7 As Integer = 16
 Public Const cmd_id_increment_value_IE8 As Integer = 105
 Public Const cmd_id_increment_value_IE9 As Integer = 105
 'Public Const cmd_id_increment_value_IE9 As Integer = 118   IE9RC'     
 '----------------------------------'
Public Sub logo_error_handler()

  Dim error_number As Integer

  error_number = tgt_inst.logo_error_code

  If error_number < 500 Then
      GoTo less_than_500
  End If
  If error_number > 499 And error_number < 999 Then
      GoTo greater_than_500
  End If
  If error_number > 999 And error_number < 2999 Then
      GoTo greater_than_1000
  End If
  If error_number > 2999 Then
      GoTo greater_than_3000
  End If
  Exit Sub
less_than_500:
  Select Case error_number
      Case 3
          tgt_inst.logo_error_text = "3 - Return without GoSub"
          tgt_inst.display_logo_error_msg()
      Case 5
          tgt_inst.logo_error_text = "5 - Invalid procedure call or argument"
          tgt_inst.display_logo_error_msg()
      Case 6
          tgt_inst.logo_error_text = "6 - Overflow Error"
          tgt_inst.display_logo_error_msg()
      Case 7
          tgt_inst.logo_error_text = "7 - Out of Memory.  The operation could not allocate enough memory."
          tgt_inst.display_logo_error_msg()
      Case 9
          tgt_inst.logo_error_text = "9 - Subscript out of Range"
          tgt_inst.display_logo_error_msg()
      Case 10
          tgt_inst.logo_error_text = "10 - This array is fixed or temporarily locked"
          tgt_inst.display_logo_error_msg()
      Case 11
          tgt_inst.logo_error_text = "11 - Division by Zero Error"
          tgt_inst.display_logo_error_msg()
      Case 13
          tgt_inst.logo_error_text = "13 - Type Mismatch error"
          tgt_inst.display_logo_error_msg()
      Case 14
          tgt_inst.logo_error_text = "14 - Out of String Space"
          tgt_inst.display_logo_error_msg()
      Case 16
          tgt_inst.logo_error_text = "16 - Expression too complex"
          tgt_inst.display_logo_error_msg()
      Case 17
          tgt_inst.logo_error_text = "17 - Can't perform requested operation"
          tgt_inst.display_logo_error_msg()
      Case 18
          tgt_inst.logo_error_text = "18 - User Interrupt Occurred"
          tgt_inst.display_logo_error_msg()
      Case 20
          tgt_inst.logo_error_text = "20 - Resume without Error"
          tgt_inst.display_logo_error_msg()
      Case 28
          tgt_inst.logo_error_text = "28 - Out of Stack Space"
         tgt_inst.display_logo_error_msg()
      Case 35
          tgt_inst.logo_error_text = "35 - Sub, Function, or Property not defined"
         tgt_inst.display_logo_error_msg()
      Case 47
          tgt_inst.logo_error_text = "47 - Too many DLL application clients"
         tgt_inst.display_logo_error_msg()
      Case 48
          tgt_inst.logo_error_text = "48 - Error in loading DLL"
         tgt_inst.display_logo_error_msg()
      Case 49
          tgt_inst.logo_error_text = "49 - Bad DLL calling convention"
         tgt_inst.display_logo_error_msg()
      Case 51
          tgt_inst.logo_error_text = "51 - Internal Error"
         tgt_inst.display_logo_error_msg()
      Case 52
          tgt_inst.logo_error_text = "52 - Bad filename or number"
         tgt_inst.display_logo_error_msg()
      Case 53
          tgt_inst.logo_error_text = "53 - File not found"
         tgt_inst.display_logo_error_msg()
      Case 54
          tgt_inst.logo_error_text = "54 - Bad File Mode (2 different access methods used!)"
         tgt_inst.display_logo_error_msg()
      Case 55
          tgt_inst.logo_error_text = "55 - File already open"
         tgt_inst.display_logo_error_msg()
      Case 57
          tgt_inst.logo_error_text = "57 - Device I/O Error"
         tgt_inst.display_logo_error_msg()
      Case 58
          tgt_inst.logo_error_text = "58 - File already exists"
         tgt_inst.display_logo_error_msg()
      Case 59
          tgt_inst.logo_error_text = "59 - Bad record length"
         tgt_inst.display_logo_error_msg()
      Case 61
          tgt_inst.logo_error_text = "61 - Disk is Full"
         tgt_inst.display_logo_error_msg()
      Case 62
          tgt_inst.logo_error_text = "62 - Input past end of file"
         tgt_inst.display_logo_error_msg()
      Case 63
          tgt_inst.logo_error_text = "63 - Bad file number"
         tgt_inst.display_logo_error_msg()
      Case 64
          tgt_inst.logo_error_text = "64 - Bad filename"
         tgt_inst.display_logo_error_msg()
      Case 67
          tgt_inst.logo_error_text = "67 - Too many filename at the same time"
         tgt_inst.display_logo_error_msg()
      Case 68
          tgt_inst.logo_error_text = "68 - Device unavailable"
         tgt_inst.display_logo_error_msg()
      Case 70
          tgt_inst.logo_error_text = "70 - Permission denied."
         tgt_inst.display_logo_error_msg()
      Case 71
          tgt_inst.logo_error_text = "71 - Disk not ready"
         tgt_inst.display_logo_error_msg()
      Case 72
          tgt_inst.logo_error_text = "72 - Disk media error"
         tgt_inst.display_logo_error_msg()
      Case 74
          tgt_inst.logo_error_text = "74 - Can't rename files across different drives"
         tgt_inst.display_logo_error_msg()
      Case 75
          tgt_inst.logo_error_text = "75 - Path/file access error"
         tgt_inst.display_logo_error_msg()
      Case 76
          tgt_inst.logo_error_text = "76 - Path not found"
         tgt_inst.display_logo_error_msg()
      Case 91
          tgt_inst.logo_error_text = "91 - Object variable or With block variable not set"
         tgt_inst.display_logo_error_msg()
      Case 92
          tgt_inst.logo_error_text = "92 - For Loop not initialized"
         tgt_inst.display_logo_error_msg()
      Case 93
          tgt_inst.logo_error_text = "93 - Invalid pattern string"
         tgt_inst.display_logo_error_msg()
      Case 94
          tgt_inst.logo_error_text = "94 - Invalid use of Null"
         tgt_inst.display_logo_error_msg()
      Case 97
          tgt_inst.logo_error_text = "97 - Can't call Friend procedure on an object that is not an" & " instance of the defining class"
         tgt_inst.display_logo_error_msg()
      Case 98
          tgt_inst.logo_error_text = "98 - A property of method call cannot include a reference to a private " & "object, either as an argument or as a return value"
         tgt_inst.display_logo_error_msg()
      Case 100
          tgt_inst.logo_error_text = "100 - JET_errRfsFailure"
         tgt_inst.display_logo_error_msg()
      Case 101
          tgt_inst.logo_error_text = "101 - JET_errRfsFailure"
         tgt_inst.display_logo_error_msg()
      Case 102
          tgt_inst.logo_error_text = "102 - Could not close DOS file"
         tgt_inst.display_logo_error_msg()
      Case 103
          tgt_inst.logo_error_text = "103 - Could not start thread"
         tgt_inst.display_logo_error_msg()
      Case 104
          tgt_inst.logo_error_text = "104 - Fail to get computername"
         tgt_inst.display_logo_error_msg()
      Case 105
          tgt_inst.logo_error_text = "105 - System busy due to too many IOs"
         tgt_inst.display_logo_error_msg()
      Case 200
          tgt_inst.logo_error_text = "200 - Buffer page evicted"
         tgt_inst.display_logo_error_msg()
      Case 201
          tgt_inst.logo_error_text = "201 - Page not found"
         tgt_inst.display_logo_error_msg()
      Case 202
          tgt_inst.logo_error_text = "202 - Cannot abandon buffer"
         tgt_inst.display_logo_error_msg()

      Case 260
          tgt_inst.logo_error_text = "260 - No timer available"
         tgt_inst.display_logo_error_msg()
          'these errors are from the MCI list of trappable errors'
      Case 277
          tgt_inst.logo_error_text = "277 - MCIErr Internal Error"
         tgt_inst.display_logo_error_msg()
      Case 278
          tgt_inst.logo_error_text = "278 - MCIErr Driver Error"
         tgt_inst.display_logo_error_msg()
      Case 279
          tgt_inst.logo_error_text = "279 - MCIErr Cannot Use All"
         tgt_inst.display_logo_error_msg()
      Case 280
          tgt_inst.logo_error_text = "280 - MCIErr Multiple Errors"
         tgt_inst.display_logo_error_msg()
      Case 281
          tgt_inst.logo_error_text = "281 - MCIErr Extension not found"
         tgt_inst.display_logo_error_msg()
      Case 282
          tgt_inst.logo_error_text = "282 - No foreign application responded to DDE initiate"
         tgt_inst.display_logo_error_msg()
      Case 285
          tgt_inst.logo_error_text = "285 - Foreign application won't perform DDE method or operation"
         tgt_inst.display_logo_error_msg()
      Case 286
          tgt_inst.logo_error_text = "286 - Timeout whilewaiting for DDE response"
         tgt_inst.display_logo_error_msg()
      Case 287
          tgt_inst.logo_error_text = "287 - User pressed excape key during DDE operation."
         tgt_inst.display_logo_error_msg()
      Case 288
          tgt_inst.logo_error_text = "288 - Destination is busy."
         tgt_inst.display_logo_error_msg()
      Case 290
          tgt_inst.logo_error_text = "290 - Data in wrong format."
         tgt_inst.display_logo_error_msg()
      Case 293
          tgt_inst.logo_error_text = "293 - DDE method invoked with no channel open."
         tgt_inst.display_logo_error_msg()
      Case 294
          tgt_inst.logo_error_text = "294 - Invalid DDE Link Format."
         tgt_inst.display_logo_error_msg()
      Case 295
          tgt_inst.logo_error_text = "295 - Message queue filled; DDE message lost."
         tgt_inst.display_logo_error_msg()
      Case 296
          tgt_inst.logo_error_text = "296 - PasteLink already performed on this control."
         tgt_inst.display_logo_error_msg()
      Case 297
          tgt_inst.logo_error_text = "297 - Can't set LinkMode; invalid LinkTopic."
         tgt_inst.display_logo_error_msg()
      Case 298
          tgt_inst.logo_error_text = "298 - System DLL could not be loaded."
         tgt_inst.display_logo_error_msg()
      Case 300
          tgt_inst.logo_error_text = "300 - Out of page space"
         tgt_inst.display_logo_error_msg()
      Case 301
          tgt_inst.logo_error_text = "301 - Itag too big"
         tgt_inst.display_logo_error_msg()
      Case 302
          tgt_inst.logo_error_text = "302 - Record deleted"
         tgt_inst.display_logo_error_msg()
      Case 303
          tgt_inst.logo_error_text = "303 - Tags used up"
         tgt_inst.display_logo_error_msg()
      Case 304
          tgt_inst.logo_error_text = "304 - Conflict in BM Clean up"
         tgt_inst.display_logo_error_msg()
      Case 305
          tgt_inst.logo_error_text = "305 - No Short Circuit Avail"
         tgt_inst.display_logo_error_msg()
      Case 306
          tgt_inst.logo_error_text = "306 - Cannot horizontally split FDP"
         tgt_inst.display_logo_error_msg()
      Case 307
          tgt_inst.logo_error_text = "307 - Cannot go up"
         tgt_inst.display_logo_error_msg()
      Case 308
          tgt_inst.logo_error_text = "308 - On an FDP Node"
         tgt_inst.display_logo_error_msg()
      Case 309
          tgt_inst.logo_error_text = "309 - May have left critical section"
         tgt_inst.display_logo_error_msg()
      Case 310
          tgt_inst.logo_error_text = "310 - Moved through empty page"
         tgt_inst.display_logo_error_msg()
      Case 311
          tgt_inst.logo_error_text = "311 - Device extent being extended"
         tgt_inst.display_logo_error_msg()
      Case 312
          tgt_inst.logo_error_text = "312 - Found Less"
         tgt_inst.display_logo_error_msg()
      Case 313
          tgt_inst.logo_error_text = "313 - Found Greater"
         tgt_inst.display_logo_error_msg()
      Case 314
          tgt_inst.logo_error_text = "314 - Son out of range"
         tgt_inst.display_logo_error_msg()
      Case 315
          tgt_inst.logo_error_text = "315 - Item out of range"
         tgt_inst.display_logo_error_msg()
      Case 316
          tgt_inst.logo_error_text = "316 - Greater than all items"
         tgt_inst.display_logo_error_msg()
      Case 317
          tgt_inst.logo_error_text = "317 - Last node of item list"
         tgt_inst.display_logo_error_msg()
      Case 318
          tgt_inst.logo_error_text = "318 - First node of item list"
         tgt_inst.display_logo_error_msg()
      Case 319
          tgt_inst.logo_error_text = "319 - Duplicated Item"
         tgt_inst.display_logo_error_msg()
      Case 320
          tgt_inst.logo_error_text = "320 - Item not there, or Cannot use character device names in specified file names"
         tgt_inst.display_logo_error_msg()
      Case 321
          tgt_inst.logo_error_text = "321 - Invalid file format."
         tgt_inst.display_logo_error_msg()
      Case 322
          tgt_inst.logo_error_text = "322 - Can't create necessary temporary file."
         tgt_inst.display_logo_error_msg()
      Case 325
          tgt_inst.logo_error_text = "325 - 'Item' is not a valid resource file."
         tgt_inst.display_logo_error_msg()
      Case 326
          tgt_inst.logo_error_text = "326 - Resource with identifier 'item' not found."
         tgt_inst.display_logo_error_msg()
      Case 327
          tgt_inst.logo_error_text = "327 - Data Value named 'item' not found."
         tgt_inst.display_logo_error_msg()
      Case 328
          tgt_inst.logo_error_text = "328 - Illegal parameter.  Can't write arrays."
         tgt_inst.display_logo_error_msg()
      Case 335
          tgt_inst.logo_error_text = "335 - Could not access system registry"
         tgt_inst.display_logo_error_msg()
      Case 336
          tgt_inst.logo_error_text = "336 - Component not correctly registered"
         tgt_inst.display_logo_error_msg()
      Case 337
          tgt_inst.logo_error_text = "337 - Component not found "
         tgt_inst.display_logo_error_msg()
      Case 338
          tgt_inst.logo_error_text = "338 - Component did not run correctly "
         tgt_inst.display_logo_error_msg()
      Case 339
          tgt_inst.logo_error_text = "339 - Object server 'item' not correctly registered."
         tgt_inst.display_logo_error_msg()
      Case 340
          tgt_inst.logo_error_text = "340 - Control array element 'item' doesn't exist."
         tgt_inst.display_logo_error_msg()
      Case 341
          tgt_inst.logo_error_text = "341 - Invalid control array index."
         tgt_inst.display_logo_error_msg()
      Case 342
          tgt_inst.logo_error_text = "342 - Not enough room to allocate control array 'item'."
         tgt_inst.display_logo_error_msg()
      Case 343
          tgt_inst.logo_error_text = "343 - Object not an array."
         tgt_inst.display_logo_error_msg()
      Case 344
          tgt_inst.logo_error_text = "344 - Must specify index for object array"
         tgt_inst.display_logo_error_msg()
      Case 345
          tgt_inst.logo_error_text = "345 - Reached limit: cannot create any more controls at this time."
         tgt_inst.display_logo_error_msg()
      Case 360
          tgt_inst.logo_error_text = "360 - Object already loaded "
         tgt_inst.display_logo_error_msg()
      Case 361
          tgt_inst.logo_error_text = "361 - Can't load or unload this object "
         tgt_inst.display_logo_error_msg()
      Case 362
          tgt_inst.logo_error_text = "362 - Can't unload controls created at design time."
         tgt_inst.display_logo_error_msg()
      Case 363
          tgt_inst.logo_error_text = "363 - Control specified not found "
         tgt_inst.display_logo_error_msg()
      Case 364
          tgt_inst.logo_error_text = "364 - Object was unloaded "
         tgt_inst.display_logo_error_msg()
      Case 365
          tgt_inst.logo_error_text = "365 - Unable to unload within this context"
         tgt_inst.display_logo_error_msg()
      Case 366
          tgt_inst.logo_error_text = "366 - No MDI form available to load."
         tgt_inst.display_logo_error_msg()
      Case 367
          tgt_inst.logo_error_text = "367 - Can't load (or register) ActiveX control: 'item'."
         tgt_inst.display_logo_error_msg()
      Case 368
          tgt_inst.logo_error_text = "368 - The specified file is out of date. This program requires a later version "
         tgt_inst.display_logo_error_msg()
      Case 369
          tgt_inst.logo_error_text = "369 - Operation not valid in a DLL"
         tgt_inst.display_logo_error_msg()
      Case 370
          tgt_inst.logo_error_text = "370 - The ActiveX designer's type information doesnotmatch whatwas saved. unable to load"
         tgt_inst.display_logo_error_msg()
      Case 371
          tgt_inst.logo_error_text = "371 - The specified object can't be used as an owner form for Show "
         tgt_inst.display_logo_error_msg()
      Case 378
          tgt_inst.logo_error_text = "378 - 'Item' cannot beset while loadin."
         tgt_inst.display_logo_error_msg()
      Case 379
          tgt_inst.logo_error_text = "379 - You can't put a default or cancel button on a property page."
         tgt_inst.display_logo_error_msg()
      Case 380
          tgt_inst.logo_error_text = "380 - Invalid property value "
         tgt_inst.display_logo_error_msg()
      Case 381
          tgt_inst.logo_error_text = "381 - Invalid property-array index "
         tgt_inst.display_logo_error_msg()
      Case 382
          tgt_inst.logo_error_text = "382 - Property Set can't be executed at run time "
         tgt_inst.display_logo_error_msg()
      Case 383
          tgt_inst.logo_error_text = "383 - Property Set can't be used with a read-only property "
         tgt_inst.display_logo_error_msg()
      Case 384
          tgt_inst.logo_error_text = "384 - A form can't be moved or sized while minimized or maximized."
         tgt_inst.display_logo_error_msg()
      Case 385
          tgt_inst.logo_error_text = "385 - Need property-array index "
         tgt_inst.display_logo_error_msg()
      Case 387
          tgt_inst.logo_error_text = "387 - Property Set not permitted "
         tgt_inst.display_logo_error_msg()
      Case 388
          tgt_inst.logo_error_text = "388 - Can't set Visible property from a parent menu."
         tgt_inst.display_logo_error_msg()
      Case 389
          tgt_inst.logo_error_text = "389 - Invalid Key"
         tgt_inst.display_logo_error_msg()
      Case 393
          tgt_inst.logo_error_text = "393 - Property Get can't be executed at run time "
         tgt_inst.display_logo_error_msg()
      Case 394
          tgt_inst.logo_error_text = "394 - Property Get can't be executed on write-only property "
         tgt_inst.display_logo_error_msg()
      Case 395
          tgt_inst.logo_error_text = "395 - 'Item' property is write-only"
         tgt_inst.display_logo_error_msg()
      Case 396
          tgt_inst.logo_error_text = "396 - 'Item' property cannot be set within a page."
         tgt_inst.display_logo_error_msg()
      Case 397
          tgt_inst.logo_error_text = "397 - Can't load, unload, or set visible property for top level menus while they" & " are merged."
         tgt_inst.display_logo_error_msg()
      Case 398
          tgt_inst.logo_error_text = "398 - Client site not available."
         tgt_inst.display_logo_error_msg()
      Case 399
          tgt_inst.logo_error_text = "399 - You can't put a default or Cancel button on a User control unless " & " its defaultcancel property is set."
         tgt_inst.display_logo_error_msg()
      Case 400
          tgt_inst.logo_error_text = "400 - Form already displayed; can't show modally "
         tgt_inst.display_logo_error_msg()
      Case 401
          tgt_inst.logo_error_text = "401 - Can't show non-modal form when modal form is displayed."
         tgt_inst.display_logo_error_msg()
      Case 402
          tgt_inst.logo_error_text = "402 - Code must close topmost modal form first "
         tgt_inst.display_logo_error_msg()
      Case 403
          tgt_inst.logo_error_text = "403 - MDI forms cannot be shown modally."
         tgt_inst.display_logo_error_msg()
      Case 404
          tgt_inst.logo_error_text = "404 - MDI child forms cannt be shown modally."
         tgt_inst.display_logo_error_msg()
      Case 405
          tgt_inst.logo_error_text = "405 - Unable to show modal form within this context."
         tgt_inst.display_logo_error_msg()
      Case 406
          tgt_inst.logo_error_text = "406 - Non-modal forms cannotbe displayed in this host application from an ActiveX DLL."
         tgt_inst.display_logo_error_msg()
      Case 419
          tgt_inst.logo_error_text = "419 - Permission to use object denied "
         tgt_inst.display_logo_error_msg()
      Case 422
          tgt_inst.logo_error_text = "422 - Property not found "
         tgt_inst.display_logo_error_msg()
      Case 423
          tgt_inst.logo_error_text = "423 - Property or method not found "
         tgt_inst.display_logo_error_msg()
      Case 424
          tgt_inst.logo_error_text = "424 - Object required"
         tgt_inst.display_logo_error_msg()
      Case 425
          tgt_inst.logo_error_text = "425 - Invalid object use "
         tgt_inst.display_logo_error_msg()
      Case 426
          tgt_inst.logo_error_text = "426 - Only one MDI form allowed."
         tgt_inst.display_logo_error_msg()
      Case 427
          tgt_inst.logo_error_text = "427 - Invalid object type; menu control required."
         tgt_inst.display_logo_error_msg()
      Case 428
          tgt_inst.logo_error_text = "428 - Popup menu must have at least one submenu."
         tgt_inst.display_logo_error_msg()
      Case 429
          tgt_inst.logo_error_text = "429 - Component can't create object or return reference to this object "
         tgt_inst.display_logo_error_msg()
      Case 430
          tgt_inst.logo_error_text = "430 - Class doesn't support Automation "
         tgt_inst.display_logo_error_msg()
      Case 432
          tgt_inst.logo_error_text = "432 - File name or class name not found during Automation operation "
         tgt_inst.display_logo_error_msg()
      Case 438
          tgt_inst.logo_error_text = "438 - Object doesn't support this property or method "
         tgt_inst.display_logo_error_msg()
      Case 440
          tgt_inst.logo_error_text = "440 - Automation error "
         tgt_inst.display_logo_error_msg()
      Case 442
          tgt_inst.logo_error_text = "442 - Connection to type library or object library for remote process has been lost "
         tgt_inst.display_logo_error_msg()
      Case 443
          tgt_inst.logo_error_text = "443 - Automation object doesn't have a default value "
         tgt_inst.display_logo_error_msg()
      Case 444
          tgt_inst.logo_error_text = "444 - Method not applicable in this context."
         tgt_inst.display_logo_error_msg()
      Case 445
          tgt_inst.logo_error_text = "445 - Object doesn't support this action "
         tgt_inst.display_logo_error_msg()
      Case 446
          tgt_inst.logo_error_text = "446 - Object doesn't support named arguments "
         tgt_inst.display_logo_error_msg()
      Case 447
          tgt_inst.logo_error_text = "447 - Object doesn't support current locale setting "
         tgt_inst.display_logo_error_msg()
      Case 448
          tgt_inst.logo_error_text = "448 - Named argument not found "
         tgt_inst.display_logo_error_msg()
      Case 449
          tgt_inst.logo_error_text = "449 - Argument not optional or invalid property assignment "
         tgt_inst.display_logo_error_msg()
      Case 450
          tgt_inst.logo_error_text = "450 - Wrong number of arguments or invalid property assignment "
         tgt_inst.display_logo_error_msg()
      Case 451
          tgt_inst.logo_error_text = "451 - Object not a collection"
         tgt_inst.display_logo_error_msg()
      Case 452
          tgt_inst.logo_error_text = "452 - Invalid ordinal "
         tgt_inst.display_logo_error_msg()
      Case 453
          tgt_inst.logo_error_text = "453 - Specified not found "
         tgt_inst.display_logo_error_msg()
      Case 454
          tgt_inst.logo_error_text = "454 - Code resource not found "
         tgt_inst.display_logo_error_msg()
      Case 455
          tgt_inst.logo_error_text = "455 - Code resource lock error "
         tgt_inst.display_logo_error_msg()
      Case 457
          tgt_inst.logo_error_text = "457 - This key is already associated with an element of this collection "
         tgt_inst.display_logo_error_msg()
      Case 458
          tgt_inst.logo_error_text = "458 - Variable uses a type not supported in Visual Basic "
         tgt_inst.display_logo_error_msg()
      Case 459
          tgt_inst.logo_error_text = "459 - This component doesn't support the set of events "
         tgt_inst.display_logo_error_msg()
      Case 460
          tgt_inst.logo_error_text = "460 - Invalid Clipboard format "
         tgt_inst.display_logo_error_msg()
      Case 461
          tgt_inst.logo_error_text = "461 - Method or data member not found "
         tgt_inst.display_logo_error_msg()
      Case 462
          tgt_inst.logo_error_text = "462 - The remote server machine does not exist or is unavailable "
         tgt_inst.display_logo_error_msg()
      Case 463
          tgt_inst.logo_error_text = "463 - Class not registered on local machine"
         tgt_inst.display_logo_error_msg()
      Case 480
          tgt_inst.logo_error_text = "480 - Can't create AutoRedraw image "
         tgt_inst.display_logo_error_msg()
      Case 481
          tgt_inst.logo_error_text = "481 - Invalid picture "
         tgt_inst.display_logo_error_msg()
      Case 482
          tgt_inst.logo_error_text = "482 - Printer error "
         tgt_inst.display_logo_error_msg()
      Case 483
          tgt_inst.logo_error_text = "483 - Printer driver does not support specified property "
         tgt_inst.display_logo_error_msg()
      Case 484
          tgt_inst.logo_error_text = "484 - Problem getting printer information from the system." & "Make sure the printer is set up correctly"
         tgt_inst.display_logo_error_msg()
      Case 485
          tgt_inst.logo_error_text = "485 - Invalid picture type "
         tgt_inst.display_logo_error_msg()
      Case 486
          tgt_inst.logo_error_text = "486 - Can't print form image to this type of printer "
         tgt_inst.display_logo_error_msg()
      Case 490
          tgt_inst.logo_error_text = "490 - Top-level or invalid menu specified as PopupMenu default"
         tgt_inst.display_logo_error_msg()
  End Select
  Exit Sub

greater_than_500:
  Select Case error_number

      Case 500
          tgt_inst.logo_error_text = "500 - Restore failed"
         tgt_inst.display_logo_error_msg()
      Case 501
          tgt_inst.logo_error_text = "501 - Log file is corrupt"
         tgt_inst.display_logo_error_msg()
      Case 502
          tgt_inst.logo_error_text = "502 - Last log record read"
         tgt_inst.display_logo_error_msg()
      Case 503
          tgt_inst.logo_error_text = "503 - No backup directory given"
         tgt_inst.display_logo_error_msg()
      Case 504
          tgt_inst.logo_error_text = "504 - The backup directory is not empty"
         tgt_inst.display_logo_error_msg()
      Case 505
          tgt_inst.logo_error_text = "505 - Backup is active already"
         tgt_inst.display_logo_error_msg()
      Case 506
          tgt_inst.logo_error_text = "506 - Fail to restore (copy) database"
         tgt_inst.display_logo_error_msg()
      Case 507
          tgt_inst.logo_error_text = "507 - No databases for restore found"
         tgt_inst.display_logo_error_msg()
      Case 508
          tgt_inst.logo_error_text = "508 - jet.log for restore is missing"
         tgt_inst.display_logo_error_msg()
      Case 509
          tgt_inst.logo_error_text = "509 - Missing the log file for check point"
         tgt_inst.display_logo_error_msg()
      Case 510
          tgt_inst.logo_error_text = "510 - Fail when writing to log file"
         tgt_inst.display_logo_error_msg()
      Case 511
          tgt_inst.logo_error_text = "511 - Fail to incremental backup for non-contiguous generation number"
         tgt_inst.display_logo_error_msg()
      Case 512
          tgt_inst.logo_error_text = "512 - Fail to make a temp directory"
         tgt_inst.display_logo_error_msg()
      Case 513
          tgt_inst.logo_error_text = "513 - Fail to clean up temp directory"
         tgt_inst.display_logo_error_msg()
      Case 514
          tgt_inst.logo_error_text = "514 - Version of log file is not compatible with Jet version"
         tgt_inst.display_logo_error_msg()
      Case 515
          tgt_inst.logo_error_text = "515 - Version of next log file is not compatible with current one"
         tgt_inst.display_logo_error_msg()
      Case 516
          tgt_inst.logo_error_text = "516 - Log is not active"
         tgt_inst.display_logo_error_msg()
      Case 517
          tgt_inst.logo_error_text = "517 - Log buffer is too small for recovery"
         tgt_inst.display_logo_error_msg()
      Case 518
          tgt_inst.logo_error_text = "518 - Retry to LGLogRec"
         tgt_inst.display_logo_error_msg()
      Case 520
          tgt_inst.logo_error_text = "520 - Can't empty Clipboard "
         tgt_inst.display_logo_error_msg()
      Case 521
          tgt_inst.logo_error_text = "521 - Can't open Clipboard "
         tgt_inst.display_logo_error_msg()
      Case 523
          tgt_inst.logo_error_text = "523 - The data binding DLL, 'item', could not be loaded."
         tgt_inst.display_logo_error_msg()
      Case 524
          tgt_inst.logo_error_text = "524 - Item."
         tgt_inst.display_logo_error_msg()
      Case 525
          tgt_inst.logo_error_text = "525 - Data Access Error."
         tgt_inst.display_logo_error_msg()
      Case 527
          tgt_inst.logo_error_text = "527 - The given bookmark was invalid."
         tgt_inst.display_logo_error_msg()
      Case 536
          tgt_inst.logo_error_text = "536 - Could not lock the database."
         tgt_inst.display_logo_error_msg()
      Case 537
          tgt_inst.logo_error_text = "537 - Could not lock the desired column."
         tgt_inst.display_logo_error_msg()
      Case 541
          tgt_inst.logo_error_text = "541 - Could not lock the database."
         tgt_inst.display_logo_error_msg()
      Case 542
          tgt_inst.logo_error_text = "542 - The row has been deleted since the update was started."
         tgt_inst.display_logo_error_msg()
      Case 545
          tgt_inst.logo_error_text = "545 - Unable to bind to field: 'item'."
         tgt_inst.display_logo_error_msg()
      Case 672
          tgt_inst.logo_error_text = "672 - Dataobject formats list may not be cleared or expanded outside of" & "the OLEstartdrag event."
         tgt_inst.display_logo_error_msg()
      Case 673
          tgt_inst.logo_error_text = "673 - Expected at least one argument."
         tgt_inst.display_logo_error_msg()
      Case 674
          tgt_inst.logo_error_text = "674 - Illegal recursive invocation of OLE drag and drop."
         tgt_inst.display_logo_error_msg()
      Case 675
          tgt_inst.logo_error_text = "675 - Non-intrinsic OLE drag and drop formats used with SetData require " & "Byte array data.  GeData may return more bytes than were given to SetData."
         tgt_inst.display_logo_error_msg()
      Case 676
          tgt_inst.logo_error_text = "676 - Requested data was not supplied to the DataObject during the OLESetData event."
         tgt_inst.display_logo_error_msg()
      Case 680
          tgt_inst.logo_error_text = "680 - Invalid Procedure Call in Safe Mode"
         tgt_inst.display_logo_error_msg()
      Case 688
          tgt_inst.logo_error_text = "688 - Failure in AsyncRead."
         tgt_inst.display_logo_error_msg()
      Case 689
          tgt_inst.logo_error_text = "689 - PropertyName parameter conflicts with the PropertyName of an AsyncRead in progress."
         tgt_inst.display_logo_error_msg()
      Case 690
          tgt_inst.logo_error_text = "690 - Can't find or load the required file urlmon.dll!"
         tgt_inst.display_logo_error_msg()
      Case 693
          tgt_inst.logo_error_text = "693 - An unknown protocol was specified in Target parameter."
         tgt_inst.display_logo_error_msg()
      Case 713
          tgt_inst.logo_error_text = "713 - Application-defined or object-defined error while generating standard reports"
         tgt_inst.display_logo_error_msg()

      Case 735
          tgt_inst.logo_error_text = "735 - Can't save file to TEMP directory "
         tgt_inst.display_logo_error_msg()
      Case 744
          tgt_inst.logo_error_text = "744 - Search text not found "
         tgt_inst.display_logo_error_msg()
      Case 746
          tgt_inst.logo_error_text = "746 - Replacements too long "
         tgt_inst.display_logo_error_msg()
          '1000 '
  End Select
  Exit Sub

greater_than_1000:
  Select Case error_number
      Case 1000
          tgt_inst.logo_error_text = "1000 - Termination in progress"
         tgt_inst.display_logo_error_msg()
      Case 1001
          tgt_inst.logo_error_text = "1001 - API not supported"
         tgt_inst.display_logo_error_msg()
      Case 1002
          tgt_inst.logo_error_text = "1002 - Invalid name"
         tgt_inst.display_logo_error_msg()
      Case 1003
          tgt_inst.logo_error_text = "1003 - Invalid API parameter"
         tgt_inst.display_logo_error_msg()
      Case 1004
          tgt_inst.logo_error_text = "1004 - Column is NULL-valued"
         tgt_inst.display_logo_error_msg()
      Case 1006
          tgt_inst.logo_error_text = "1006 - Buffer too small for data"
         tgt_inst.display_logo_error_msg()
      Case 1007
          tgt_inst.logo_error_text = "1007 - Database is already attached"
         tgt_inst.display_logo_error_msg()
      Case 1008
          tgt_inst.logo_error_text = "1008 - Attach a readonly database file for read/write operations"
         tgt_inst.display_logo_error_msg()
      Case 1009
          tgt_inst.logo_error_text = "1009 - Sort does not fit in memory"
         tgt_inst.display_logo_error_msg()
      Case 1010
          tgt_inst.logo_error_text = "1010 - Invalid database id"
         tgt_inst.display_logo_error_msg()
      Case 1011
          tgt_inst.logo_error_text = "1011 - Out of Memory"
         tgt_inst.display_logo_error_msg()
      Case 1012
          tgt_inst.logo_error_text = "1012 - Maximum database size reached"
         tgt_inst.display_logo_error_msg()
      Case 1013
          tgt_inst.logo_error_text = "1013 - Out of table cursors"
         tgt_inst.display_logo_error_msg()
      Case 1014
          tgt_inst.logo_error_text = "1014 - Out of database page buffers"
         tgt_inst.display_logo_error_msg()
      Case 1015
          tgt_inst.logo_error_text = "1015 - Too many indexes"
         tgt_inst.display_logo_error_msg()
      Case 1016
          tgt_inst.logo_error_text = "1016 - Too many columns in an index"
         tgt_inst.display_logo_error_msg()
      Case 1017
          tgt_inst.logo_error_text = "1017 - Record has been deleted"
         tgt_inst.display_logo_error_msg()
      Case 1018
          tgt_inst.logo_error_text = "1018 - Read verification error"
         tgt_inst.display_logo_error_msg()
      Case 1020
          tgt_inst.logo_error_text = "1020 - Out of file handles"
         tgt_inst.display_logo_error_msg()
      Case 1022
          tgt_inst.logo_error_text = "1022 - Disk IO error"
         tgt_inst.display_logo_error_msg()
      Case 1023
          tgt_inst.logo_error_text = "1023 - Invalid file path"
         tgt_inst.display_logo_error_msg()
      Case 1026
          tgt_inst.logo_error_text = "1026 - Record larger than maximum size"
         tgt_inst.display_logo_error_msg()
      Case 1027
          tgt_inst.logo_error_text = "1027 - Too many open databases"
         tgt_inst.display_logo_error_msg()
      Case 1028
          tgt_inst.logo_error_text = "1028 - Not a database file"
         tgt_inst.display_logo_error_msg()
      Case 1029
          tgt_inst.logo_error_text = "1029 - JetInit not yet called"
         tgt_inst.display_logo_error_msg()
      Case 1030
          tgt_inst.logo_error_text = "1030 - JetInit already called"
         tgt_inst.display_logo_error_msg()
      Case 1032
          tgt_inst.logo_error_text = "1032 - Cannot access file"
         tgt_inst.display_logo_error_msg()
      Case 1034
          tgt_inst.logo_error_text = "1034 - Query support unavailable"
         tgt_inst.display_logo_error_msg()
      Case 1035
          tgt_inst.logo_error_text = "1035 - SQL Link support unavailable"
         tgt_inst.display_logo_error_msg()
      Case 1038
          tgt_inst.logo_error_text = "1038 - Buffer is too small"
         tgt_inst.display_logo_error_msg()
      Case 1039
          tgt_inst.logo_error_text = "1039 - SeekLE or SeekGE didn't find exact match"
         tgt_inst.display_logo_error_msg()
      Case 1040
          tgt_inst.logo_error_text = "1040 - Too many columns defined"
         tgt_inst.display_logo_error_msg()
      Case 1043
          tgt_inst.logo_error_text = "1043 - Container is not empty"
         tgt_inst.display_logo_error_msg()
      Case 1044
          tgt_inst.logo_error_text = "1044 - Filename is invalid"
         tgt_inst.display_logo_error_msg()
      Case 1045
          tgt_inst.logo_error_text = "1045 - Invalid bookmark"
         tgt_inst.display_logo_error_msg()
      Case 1046
          tgt_inst.logo_error_text = "1046 - Column used in an index"
         tgt_inst.display_logo_error_msg()
      Case 1047
          tgt_inst.logo_error_text = "1047 - Data buffer doesn't match column size"
         tgt_inst.display_logo_error_msg()
      Case 1048
          tgt_inst.logo_error_text = "1048 - Cannot set column value"
         tgt_inst.display_logo_error_msg()
      Case 1051
          tgt_inst.logo_error_text = "1051 - Index is in use"
         tgt_inst.display_logo_error_msg()
      Case 1052
          tgt_inst.logo_error_text = "1052 - Link support unavailable"
         tgt_inst.display_logo_error_msg()
      Case 1053
          tgt_inst.logo_error_text = "1053 - Null keys are disallowed on index"
         tgt_inst.display_logo_error_msg()
      Case 1054
          tgt_inst.logo_error_text = "1054 - Operation must be within a transaction"
         tgt_inst.display_logo_error_msg()
      Case 1055
          tgt_inst.logo_error_text = "1055 - No extended error information"
         tgt_inst.display_logo_error_msg()
      Case 1058
          tgt_inst.logo_error_text = "1058 - No idle activity occured"
         tgt_inst.display_logo_error_msg()
      Case 1059
          tgt_inst.logo_error_text = "1059 -  Too many active database users"
         tgt_inst.display_logo_error_msg()
      Case 1060
          tgt_inst.logo_error_text = "1060 - Cannot append long value"
         tgt_inst.display_logo_error_msg()
      Case 1061
          tgt_inst.logo_error_text = "1061 - Invalid or unknown country code"
         tgt_inst.display_logo_error_msg()
      Case 1062
          tgt_inst.logo_error_text = "1062 - Invalid or unknown language id"
         tgt_inst.display_logo_error_msg()
      Case 1063
          tgt_inst.logo_error_text = "1063 - Invalid or unknown code page"
         tgt_inst.display_logo_error_msg()
      Case 1067
          tgt_inst.logo_error_text = "1067 - No write lock at transaction level 0"
         tgt_inst.display_logo_error_msg()
      Case 1068
          tgt_inst.logo_error_text = "1068 - Column set to NULL-value"
         tgt_inst.display_logo_error_msg()
      Case 1069
          tgt_inst.logo_error_text = "1069 - LMaxVerPages exceeded (XJET only)"
         tgt_inst.display_logo_error_msg()
      Case 1070
          tgt_inst.logo_error_text = "1070 - LCSRPerfFUCB * lMaxCursors exceeded (XJET only)"
         tgt_inst.display_logo_error_msg()
      Case 1101
          tgt_inst.logo_error_text = "1101 - Out of sessions"
         tgt_inst.display_logo_error_msg()
      Case 1102
          tgt_inst.logo_error_text = "1102 - Write lock failed due to outstanding write lock"
         tgt_inst.display_logo_error_msg()
      Case 1103
          tgt_inst.logo_error_text = "1103 - Xactions nested too deeply"
         tgt_inst.display_logo_error_msg()
      Case 1104
          tgt_inst.logo_error_text = "1104 - Invalid session handle"
         tgt_inst.display_logo_error_msg()
      Case 1107
          tgt_inst.logo_error_text = "1107 - Another session has private version of page"
         tgt_inst.display_logo_error_msg()
      Case 1108
          tgt_inst.logo_error_text = "1108 - Operation not allowed within a transaction"
         tgt_inst.display_logo_error_msg()
      Case 1201
          tgt_inst.logo_error_text = "1201 - Database already exists"
         tgt_inst.display_logo_error_msg()
      Case 1202
          tgt_inst.logo_error_text = "1202 - Database in use"
         tgt_inst.display_logo_error_msg()
      Case 1203
          tgt_inst.logo_error_text = "1203 - No such database"
         tgt_inst.display_logo_error_msg()
      Case 1204
          tgt_inst.logo_error_text = "1204 - Invalid database name"
         tgt_inst.display_logo_error_msg()
      Case 1205
          tgt_inst.logo_error_text = "1205 - Invalid number of pages"
         tgt_inst.display_logo_error_msg()
      Case 1206
          tgt_inst.logo_error_text = "1206 - Non-db file or corrupted db"
         tgt_inst.display_logo_error_msg()
      Case 1207
          tgt_inst.logo_error_text = "1207 - Database exclusively locked"
         tgt_inst.display_logo_error_msg()
      Case 1208
          tgt_inst.logo_error_text = "1208 - Cannot disable versioning for this database"
         tgt_inst.display_logo_error_msg()
      Case 1301
          tgt_inst.logo_error_text = "1301 - Open an empty table"
         tgt_inst.display_logo_error_msg()
      Case 1302
          tgt_inst.logo_error_text = "1302 - Table is exclusively locked"
         tgt_inst.display_logo_error_msg()
      Case 1303
          tgt_inst.logo_error_text = "1303 - Table already exists"
         tgt_inst.display_logo_error_msg()
      Case 1304
          tgt_inst.logo_error_text = "1304 - Table is in use, cannot lock"
         tgt_inst.display_logo_error_msg()
      Case 1305
          tgt_inst.logo_error_text = "1305 - No such table or object"
         tgt_inst.display_logo_error_msg()
      Case 1307
          tgt_inst.logo_error_text = "1307 - Bad file/index density"
         tgt_inst.display_logo_error_msg()
      Case 1308
          tgt_inst.logo_error_text = "1308 - Cannot define clustered index"
         tgt_inst.display_logo_error_msg()
      Case 1310
          tgt_inst.logo_error_text = "1310 - Invalid table id"
         tgt_inst.display_logo_error_msg()
      Case 1311
          tgt_inst.logo_error_text = "1311 - Cannot open any more tables"
         tgt_inst.display_logo_error_msg()
      Case 1312
          tgt_inst.logo_error_text = "1312 - Oper. not supported on table"
         tgt_inst.display_logo_error_msg()
      Case 1314
          tgt_inst.logo_error_text = "1314 - Table or object name in use"
         tgt_inst.display_logo_error_msg()
      Case 1316
          tgt_inst.logo_error_text = "1316 - Object is invalid for operation"
         tgt_inst.display_logo_error_msg()
      Case 1401
          tgt_inst.logo_error_text = "1401 - Cannot build clustered index"
         tgt_inst.display_logo_error_msg()
      Case 1402
          tgt_inst.logo_error_text = "1402 - Primary index already defined"
         tgt_inst.display_logo_error_msg()
      Case 1403
          tgt_inst.logo_error_text = "1403 - Index is already defined"
         tgt_inst.display_logo_error_msg()
      Case 1404
          tgt_inst.logo_error_text = "1404 - No such index"
         tgt_inst.display_logo_error_msg()
      Case 1405
          tgt_inst.logo_error_text = "1405 - Cannot delete clustered index"
         tgt_inst.display_logo_error_msg()
      Case 1406
          tgt_inst.logo_error_text = "1406 - Illegal index definition"
         tgt_inst.display_logo_error_msg()
      Case 1408
          tgt_inst.logo_error_text = "1408 - Clustered index already defined"
         tgt_inst.display_logo_error_msg()
      Case 1409
          tgt_inst.logo_error_text = "1409 - Invalid create index description"
         tgt_inst.display_logo_error_msg()
      Case 1410
          tgt_inst.logo_error_text = "1410 - Out of index description blocks"
         tgt_inst.display_logo_error_msg()
      Case 1501
          tgt_inst.logo_error_text = "1501 - Column value is long"
         tgt_inst.display_logo_error_msg()
      Case 1502
          tgt_inst.logo_error_text = "1502 - No such chunk in long value"
         tgt_inst.display_logo_error_msg()
      Case 1503
          tgt_inst.logo_error_text = "1503 - Field will not fit in record"
         tgt_inst.display_logo_error_msg()
      Case 1504
          tgt_inst.logo_error_text = "1504 - Null not valid"
         tgt_inst.display_logo_error_msg()
      Case 1505
          tgt_inst.logo_error_text = "1505 - Column indexed, cannot delete"
         tgt_inst.display_logo_error_msg()
      Case 1506
          tgt_inst.logo_error_text = "1506 - Field length is > maximum"
         tgt_inst.display_logo_error_msg()
      Case 1507
          tgt_inst.logo_error_text = "1507 - No such column"
         tgt_inst.display_logo_error_msg()
      Case 1508
          tgt_inst.logo_error_text = "1508 - Field is already defined"
         tgt_inst.display_logo_error_msg()
      Case 1510
          tgt_inst.logo_error_text = "1510 - Second autoinc or version column"
         tgt_inst.display_logo_error_msg()
      Case 1511
          tgt_inst.logo_error_text = "1511 - Invalid column data type"
         tgt_inst.display_logo_error_msg()
      Case 1512
          tgt_inst.logo_error_text = "1512 - Max length too big, truncated"
         tgt_inst.display_logo_error_msg()
      Case 1513
          tgt_inst.logo_error_text = "1513 - Cannot index Bit,LongText,LongBinary"
         tgt_inst.display_logo_error_msg()
      Case 1514
          tgt_inst.logo_error_text = "1514 - No non-NULL tagged columns"
         tgt_inst.display_logo_error_msg()
      Case 1515
          tgt_inst.logo_error_text = "1515 - Invalid w/o a current index"
         tgt_inst.display_logo_error_msg()
      Case 1516
          tgt_inst.logo_error_text = "1516 - The key is completely made"
         tgt_inst.display_logo_error_msg()
      Case 1517
          tgt_inst.logo_error_text = "1517 - Column Id Incorrect"
         tgt_inst.display_logo_error_msg()
      Case 1518
          tgt_inst.logo_error_text = "1518 - Bad itagSequence for tagged column"
         tgt_inst.display_logo_error_msg()
      Case 1519
          tgt_inst.logo_error_text = "1519 - Cannot delete, column participates in relationship"
         tgt_inst.display_logo_error_msg()
      Case 1520
          tgt_inst.logo_error_text = "1520 - Single instance column bursted"
         tgt_inst.display_logo_error_msg()
      Case 1521
          tgt_inst.logo_error_text = "1521 - AutoIncrement and Version cannot be tagged"
         tgt_inst.display_logo_error_msg()
      Case 1601
          tgt_inst.logo_error_text = "1601 - The key was not found"
         tgt_inst.display_logo_error_msg()
      Case 1602
          tgt_inst.logo_error_text = "1602 - No working buffer"
         tgt_inst.display_logo_error_msg()
      Case 1603
          tgt_inst.logo_error_text = "1603 - Currency not on a record"
         tgt_inst.display_logo_error_msg()
      Case 1604
          tgt_inst.logo_error_text = "1604 - Clustered key may not change"
         tgt_inst.display_logo_error_msg()
      Case 1605
          tgt_inst.logo_error_text = "1605 - Illegal duplicate key"
         tgt_inst.display_logo_error_msg()
      Case 1607
          tgt_inst.logo_error_text = "1607 - Already copy/clear current"
         tgt_inst.display_logo_error_msg()
      Case 1608
          tgt_inst.logo_error_text = "1608 -  No call to JetMakeKey"
         tgt_inst.display_logo_error_msg()
      Case 1609
          tgt_inst.logo_error_text = "1609 - No call to JetPrepareUpdate"
         tgt_inst.display_logo_error_msg()
      Case 1610
          tgt_inst.logo_error_text = "1610 - Data has changed"
         tgt_inst.display_logo_error_msg()
      Case 1611
          tgt_inst.logo_error_text = "1611 - Data has changed, operation aborted"
         tgt_inst.display_logo_error_msg()
      Case 1618
          tgt_inst.logo_error_text = "1618 - Moved to new key"
         tgt_inst.display_logo_error_msg()
      Case 1701
          tgt_inst.logo_error_text = "1701 - Too many sort processes"
         tgt_inst.display_logo_error_msg()
      Case 1702
          tgt_inst.logo_error_text = "1702 - Invalid operation on Sort"
         tgt_inst.display_logo_error_msg()
      Case 1803
          tgt_inst.logo_error_text = "1803 - Temp file could not be opened"
         tgt_inst.display_logo_error_msg()
      Case 1805
          tgt_inst.logo_error_text = "1805 - Too many open databases"
         tgt_inst.display_logo_error_msg()
      Case 1808
          tgt_inst.logo_error_text = "1808 - No space left on disk"
         tgt_inst.display_logo_error_msg()
      Case 1809
          tgt_inst.logo_error_text = "1809 - Permission denied"
         tgt_inst.display_logo_error_msg()
      Case 1811
          tgt_inst.logo_error_text = "1811 - File not found"
         tgt_inst.display_logo_error_msg()
      Case 1813
          tgt_inst.logo_error_text = "1813 - Database file is read only"
         tgt_inst.display_logo_error_msg()
      Case 1850
          tgt_inst.logo_error_text = "1850 - Cannot Restore after init."
         tgt_inst.display_logo_error_msg()
      Case 1852
          tgt_inst.logo_error_text = "1852 - Logs could not be interpreted"
         tgt_inst.display_logo_error_msg()
      Case 1906
          tgt_inst.logo_error_text = "1906 - Invalid operation"
         tgt_inst.display_logo_error_msg()
      Case 1907
          tgt_inst.logo_error_text = "1907 - Access denied"
         tgt_inst.display_logo_error_msg()
      Case 1908
          tgt_inst.logo_error_text = "1908 - Idle registry full"
         tgt_inst.display_logo_error_msg()
      Case 2001
          tgt_inst.logo_error_text = "2001 - Cannot Open File"
         tgt_inst.display_logo_error_msg()
      Case 2007
          tgt_inst.logo_error_text = "2007 - Invalid Bands Collection Index or Error 2001 Cannot Open File"
         tgt_inst.display_logo_error_msg()

  End Select
  Exit Sub

        '3000 and over codes'
greater_than_3000:
  Select Case error_number
      Case 3011
          tgt_inst.logo_error_text = "MS jet could not find specified object.  Make sure the object exists and that you spell its name and the path name correctly"
         tgt_inst.display_logo_error_msg()
      Case 3015
          tgt_inst.logo_error_text = "3015 - databasename.mdb isnt an index in this table. Look in the Indexes collection of the TableDef object to determine the valid index names. "
         tgt_inst.display_logo_error_msg()
      Case 3021
          tgt_inst.logo_error_text = "3021 - record not found - table may be empty"
         tgt_inst.display_logo_error_msg()
      Case 3022
          tgt_inst.logo_error_text = "3022 - record already exists"
         tgt_inst.display_logo_error_msg()
      Case 3024
          tgt_inst.logo_error_text = "3024 - database not found"
         tgt_inst.display_logo_error_msg()
      Case 3028
          tgt_inst.logo_error_text = "3028 - Can't Start the Application"
         tgt_inst.display_logo_error_msg()
      Case 3029
          tgt_inst.logo_error_text = "3029 - Invalid Account Name or Password"
         tgt_inst.display_logo_error_msg()

      Case 3043
          tgt_inst.logo_error_text = "3043 - Disk or Network Error"
         tgt_inst.display_logo_error_msg()
      Case 3050
          tgt_inst.logo_error_text = "3050 - Couldn't Lock File"
         tgt_inst.display_logo_error_msg()
      Case 3051
          tgt_inst.logo_error_text = "3051 - table locked - another user may be using table"
         tgt_inst.display_logo_error_msg()
      Case 3061
          tgt_inst.logo_error_text = "3061 - too few parameters.  Expected 1.'"
         tgt_inst.display_logo_error_msg()
      Case 3070
          tgt_inst.logo_error_text = "3070 - MS Jet database engine does not recognize a field as a valid field name or expression'"
         tgt_inst.display_logo_error_msg()
      Case 3075
          tgt_inst.logo_error_text = "3075 - Syntax error (missing operator) in query expression"
         tgt_inst.display_logo_error_msg()
      Case 3078
          tgt_inst.logo_error_text = "3078 - The MS Jet database engine could not find the input table or query. " & "Make sure it exists and that its name is spelled correctly"
         tgt_inst.display_logo_error_msg()
      Case 3112
          tgt_inst.logo_error_text = "3112 - Records could not be read; no read permission on the database file"
         tgt_inst.display_logo_error_msg()
      Case 3159
          tgt_inst.logo_error_text = "3159 - Not a valid bookmark"
         tgt_inst.display_logo_error_msg()
      Case 3167
          tgt_inst.logo_error_text = "3167 - Record is deleted or Error Number 3075 exists " & "Syntax error (missing operator) in query expression."
         tgt_inst.display_logo_error_msg()
      Case 3170
          tgt_inst.logo_error_text = "Could not find installable ISAM"
         tgt_inst.display_logo_error_msg()
      Case 3197
          tgt_inst.logo_error_text = "3197 - The Microsoft Jet database engine stopped the process because you and another user " & "are attempting to change the same data at the same time. "
         tgt_inst.display_logo_error_msg()
      Case 3315
          tgt_inst.logo_error_text = "3315 - Duplicate Key in Table!"
         tgt_inst.display_logo_error_msg()
      Case 3343
          tgt_inst.logo_error_text = "3343 - Unrecognized database format 'databasename.mdb'. "
         tgt_inst.display_logo_error_msg()
      Case 3464
          tgt_inst.logo_error_text = "3464 - Data type mismatch in criteria expression"
         tgt_inst.display_logo_error_msg()
      Case 3709
          tgt_inst.logo_error_text = "3709 - Web Module Error"
         tgt_inst.display_logo_error_msg()
  End Select
  Exit Sub

greater_than_5000:
  Select Case error_number
      Case 20476
          tgt_inst.logo_error_text = "20476 - The filename buffer is too small to store the selected file name."
         tgt_inst.display_logo_error_msg()
      Case 20477
          tgt_inst.logo_error_text = "20477 - Invalid Filename"
         tgt_inst.display_logo_error_msg()
      Case 20478
          tgt_inst.logo_error_text = "20478 - An attempt to subclass a listbox failed due to insufficient memory."
         tgt_inst.display_logo_error_msg()
      Case 20500
          tgt_inst.logo_error_text = "20500 - Not Enough memory for operation"
         tgt_inst.display_logo_error_msg()
      Case 20533
          tgt_inst.logo_error_text = "20533 - Unable to Open Database"
         tgt_inst.display_logo_error_msg()
      Case 24010
          tgt_inst.logo_error_text = "24010 - An unexpected internal error has occurred"
         tgt_inst.display_logo_error_msg()
      Case 24011
          tgt_inst.logo_error_text = "24011 - An error has occurred with no information available"
         tgt_inst.display_logo_error_msg()
      Case 24012
          tgt_inst.logo_error_text = "24012 - An error has occurred. Unable to retrieve error information"
         tgt_inst.display_logo_error_msg()
      Case 24013
          tgt_inst.logo_error_text = "24013 - A control cancelled the operation or an unexpected internal error has occurred"
         tgt_inst.display_logo_error_msg()

      Case 24014
          tgt_inst.logo_error_text = "24014 - Could not refresh controls"
         tgt_inst.display_logo_error_msg()
      Case 24015
          tgt_inst.logo_error_text = "24015 - Invalid property value"
         tgt_inst.display_logo_error_msg()
      Case 24016
          tgt_inst.logo_error_text = "24016 - Invalid object"
         tgt_inst.display_logo_error_msg()
      Case 24017
          tgt_inst.logo_error_text = "24017 - Method cannot be called in RDC's current state"
         tgt_inst.display_logo_error_msg()
      Case 24018
          tgt_inst.logo_error_text = "24018 - One or more of the arguments is invalid"
         tgt_inst.display_logo_error_msg()
      Case 24019
          tgt_inst.logo_error_text = "24019 - Resultset is empty"
         tgt_inst.display_logo_error_msg()
      Case 24020
          tgt_inst.logo_error_text = "24020 - Out of memory"
         tgt_inst.display_logo_error_msg()
      Case 24021
          tgt_inst.logo_error_text = "24021 - Resultset not available."
         tgt_inst.display_logo_error_msg()
      Case 24022
          tgt_inst.logo_error_text = "24022 - The connection is not open"
         tgt_inst.display_logo_error_msg()
      Case 24023
          tgt_inst.logo_error_text = "24023 - Property cannot be set in RDC's current state"
         tgt_inst.display_logo_error_msg()
      Case 24024
          tgt_inst.logo_error_text = "24024 - Property not available in RDC's current state"
         tgt_inst.display_logo_error_msg()
      Case 24025
          tgt_inst.logo_error_text = "24025 - Type Mismatch error"
         tgt_inst.display_logo_error_msg()
      Case 24026
          tgt_inst.logo_error_text = "24026 - Cannot connect to RemoteData Objects"
         tgt_inst.display_logo_error_msg()
      Case 24027
          tgt_inst.logo_error_text = "24027 - Data Type Conversion Error"
         tgt_inst.display_logo_error_msg()

      Case 24574
          tgt_inst.logo_error_text = "24574 - No fonts exist."
         tgt_inst.display_logo_error_msg()
      Case 28660
          tgt_inst.logo_error_text = "28660 - The devices section of the file win.ini did not contain an entry for the requested " & " printer."
         tgt_inst.display_logo_error_msg()
      Case 28661
          tgt_inst.logo_error_text = "28661 - The PrintDlg function failed when it attempted to create an information context."
         tgt_inst.display_logo_error_msg()
      Case 28662
          tgt_inst.logo_error_text = "28662 - The data in the DEVMODE and DEVNAMES data structures describes two different printers."
         tgt_inst.display_logo_error_msg()
      Case 28663
          tgt_inst.logo_error_text = "28663 - A default printer does not exist."
         tgt_inst.display_logo_error_msg()
      Case 28664
          tgt_inst.logo_error_text = "28664 - No printer device drivers were found."
         tgt_inst.display_logo_error_msg()
      Case 28665
          tgt_inst.logo_error_text = "28665 - The PrintDlg function failed during initialization."
         tgt_inst.display_logo_error_msg()
      Case 28666
          tgt_inst.logo_error_text = "28666 - The printer device driver failed to initialize a DEVMODE data structure."
         tgt_inst.display_logo_error_msg()
      Case 28667
          tgt_inst.logo_error_text = "28667 - The PrintDlg function failed to load the specified printer's device driver."
         tgt_inst.display_logo_error_msg()
      Case 28668
          tgt_inst.logo_error_text = "28668 - The PD_RETURNDEFAULT flag was set in the Flags member of the PRINTDLG data structure " & "but either the hDevMode or hDevNames field were nonzero."
         tgt_inst.display_logo_error_msg()
      Case 28669
          tgt_inst.logo_error_text = "28669 - The common dialog function failed to parse the strings in the [devices] section of the " & " file WIN.INI "
         tgt_inst.display_logo_error_msg()
      Case 28670
          tgt_inst.logo_error_text = "28670 - Load of required resources failed."
         tgt_inst.display_logo_error_msg()
      Case 28671
          tgt_inst.logo_error_text = "28671 - The PD_RETURNDEFAULT flag was set in the Flags member of the PRINTDLG data structure, " & "but either the hDevMode or hDevNames field were nonzero."
         tgt_inst.display_logo_error_msg()
      Case 31001
          tgt_inst.logo_error_text = "31001 - Out of memory "
         tgt_inst.display_logo_error_msg()
      Case 31003
          tgt_inst.logo_error_text = "31003 - Can't open Clipboard."
         tgt_inst.display_logo_error_msg()
      Case 31004
          tgt_inst.logo_error_text = "31004 - No object "
         tgt_inst.display_logo_error_msg()
      Case 31006
          tgt_inst.logo_error_text = "31006 - Unable to close object."
         tgt_inst.display_logo_error_msg()
      Case 31007
          tgt_inst.logo_error_text = "31007 - Can't paste."
         tgt_inst.display_logo_error_msg()
      Case 31008
          tgt_inst.logo_error_text = "31008 - Invalid property Value."
         tgt_inst.display_logo_error_msg()
      Case 31009
          tgt_inst.logo_error_text = "31009 - Can't Copy."
         tgt_inst.display_logo_error_msg()
      Case 31017
          tgt_inst.logo_error_text = "31017 - Invalid Format."
         tgt_inst.display_logo_error_msg()
      Case 31018
          tgt_inst.logo_error_text = "31018 - Class is not set "
         tgt_inst.display_logo_error_msg()
      Case 31019
          tgt_inst.logo_error_text = "31019 - Source document is not set."
         tgt_inst.display_logo_error_msg()
      Case 31021
          tgt_inst.logo_error_text = "31021 - Invalid Action."
         tgt_inst.display_logo_error_msg()
      Case 31023
          tgt_inst.logo_error_text = "31023 - Invalid or unknown Class."
         tgt_inst.display_logo_error_msg()
      Case 31024
          tgt_inst.logo_error_text = "31024 - Unable to create link."
         tgt_inst.display_logo_error_msg()
      Case 31026
          tgt_inst.logo_error_text = "31026 - Source name is too long."
         tgt_inst.display_logo_error_msg()
      Case 31027
          tgt_inst.logo_error_text = "31027 - Unable to activate object "
         tgt_inst.display_logo_error_msg()
      Case 31028
          tgt_inst.logo_error_text = "31028 - Object not running."
         tgt_inst.display_logo_error_msg()
      Case 31029
          tgt_inst.logo_error_text = "31029 - Dialog already in use."
         tgt_inst.display_logo_error_msg()
      Case 31031
          tgt_inst.logo_error_text = "31031 - Invalid source for link."
         tgt_inst.display_logo_error_msg()
      Case 31032
          tgt_inst.logo_error_text = "31032 - Unable to create embedded object "
         tgt_inst.display_logo_error_msg()
      Case 31033
          tgt_inst.logo_error_text = "31033 - Unable to fetch link source name."
         tgt_inst.display_logo_error_msg()
      Case 31034
          tgt_inst.logo_error_text = "31034 - Invalid Verb Index."
         tgt_inst.display_logo_error_msg()
      Case 31035
          tgt_inst.logo_error_text = "31035 - Incorrect Clipboard format."
         tgt_inst.display_logo_error_msg()
      Case 31036
          tgt_inst.logo_error_text = "31036 - Error saving to file "
         tgt_inst.display_logo_error_msg()
      Case 31037
          tgt_inst.logo_error_text = "31037 - Error loading from file "
         tgt_inst.display_logo_error_msg()
      Case 31039
          tgt_inst.logo_error_text = "31039 - Unable to access source document."
         tgt_inst.display_logo_error_msg()
      Case 31040
          tgt_inst.logo_error_text = "31040 - You cannot set DisplayType while the control contains an object."
         tgt_inst.display_logo_error_msg()
      Case 31041
          tgt_inst.logo_error_text = "31041 - Cannot create embedded object.  OleTypeAllowed property of 'item' " & " control is set to linked."
         tgt_inst.display_logo_error_msg()
      Case 32000
          tgt_inst.logo_error_text = "32000 - Picture format not supported"
         tgt_inst.display_logo_error_msg()
      Case 32001
          tgt_inst.logo_error_text = "32001 - Unable to obtain memory device context"
         tgt_inst.display_logo_error_msg()
      Case 32002
          tgt_inst.logo_error_text = "32002 - Unspecified Failure has occurred."
         tgt_inst.display_logo_error_msg()
      Case 32003
          tgt_inst.logo_error_text = "32003 - Unable to obtain bitmap"
         tgt_inst.display_logo_error_msg()
      Case 32004
          tgt_inst.logo_error_text = "32004 - Unable to select bitmap object"
         tgt_inst.display_logo_error_msg()
      Case 32005
          tgt_inst.logo_error_text = "32005 - Unable to allocate internal picture structure"
         tgt_inst.display_logo_error_msg()
      Case 32006
          tgt_inst.logo_error_text = "32006 - Bad GraphicCell Index"
         tgt_inst.display_logo_error_msg()
      Case 32007
          tgt_inst.logo_error_text = "32007 - No GraphicCell picture size specified."
         tgt_inst.display_logo_error_msg()
      Case 32008
          tgt_inst.logo_error_text = "32008 - Only bitmap GraphicCell pictures allowed."
         tgt_inst.display_logo_error_msg()
      Case 32010
          tgt_inst.logo_error_text = "32010 - Bad GraphicCell PictureClip property request."
         tgt_inst.display_logo_error_msg()
      Case 32012
          tgt_inst.logo_error_text = "32012 - GetObject() Windows function failure"
         tgt_inst.display_logo_error_msg()
      Case 32014
          tgt_inst.logo_error_text = "32014 - Unknown Recipient or Ambiguous Recipient while sending email"
         tgt_inst.display_logo_error_msg()
      Case 32015
          tgt_inst.logo_error_text = "32015 - Clip region boundary error."
         tgt_inst.display_logo_error_msg()
      Case 32016
          tgt_inst.logo_error_text = "32016 - Cell size too small (must be at least 1 by 1 pixel)"
         tgt_inst.display_logo_error_msg()
      Case 32017
          tgt_inst.logo_error_text = "32017 - Rows property must be greater than zero"
         tgt_inst.display_logo_error_msg()
      Case 32018
          tgt_inst.logo_error_text = "32018 - Cols property must be greater than zero"
         tgt_inst.display_logo_error_msg()
      Case 32019
          tgt_inst.logo_error_text = "32019 - StretchX property cannot be negative"
         tgt_inst.display_logo_error_msg()
      Case 32020
          tgt_inst.logo_error_text = "32020 - StretchY property cannot be negative"
         tgt_inst.display_logo_error_msg()
      Case 32021
          tgt_inst.logo_error_text = "32021 - Unknown Recipient or Ambiguous Recipient while sending email"
         tgt_inst.display_logo_error_msg()
      Case 32751
          tgt_inst.logo_error_text = "32751 - Help Call Fail.  Check Help Properties."
         tgt_inst.display_logo_error_msg()
      Case 32752
          tgt_inst.logo_error_text = "32752 - Low on memory!  Can't bring up the dialog."
         tgt_inst.display_logo_error_msg()
      Case 32753
          tgt_inst.logo_error_text = "32753 - Couldn't determine procedure address(es).  Select a different DLL."
         tgt_inst.display_logo_error_msg()
      Case 32754
          tgt_inst.logo_error_text = "32754 - DLL selected couldn't be loaded."
         tgt_inst.display_logo_error_msg()
      Case 32755
          tgt_inst.logo_error_text = "32755 - Cancel was selected."
         tgt_inst.display_logo_error_msg()
      Case 32756
          tgt_inst.logo_error_text = "32756 - The ENABLEHOOK flag was set in the Flags member of a common dialog " & "data structure but the application failed to provide a pointer to a corresponding " & " hook function."
         tgt_inst.display_logo_error_msg()
      Case 32757
          tgt_inst.logo_error_text = "32757 - The common dialog function was unable to lock the memory associated with a handle."
         tgt_inst.display_logo_error_msg()
      Case 32758
          tgt_inst.logo_error_text = "32758 - The common dialog function was unable to allocate memory for internal data structures."
         tgt_inst.display_logo_error_msg()
      Case 32759
          tgt_inst.logo_error_text = "32759 - The common dialog function failed to lock a specified resource."
         tgt_inst.display_logo_error_msg()
      Case 32760
          tgt_inst.logo_error_text = "32760 - The common dialog function failed to load a specified resource."
         tgt_inst.display_logo_error_msg()
      Case 32761
          tgt_inst.logo_error_text = "32761 - The common dialog function failed to find a specified resource."
         tgt_inst.display_logo_error_msg()
      Case 32762
          tgt_inst.logo_error_text = "32762 - The common dialog function failed to load a specified string."
         tgt_inst.display_logo_error_msg()
      Case 32763
          tgt_inst.logo_error_text = "32763 - The ENABLETEMPLATE flag was set in the Flags member of a common dialog data structure " & "but the application failed to provide a corresponding instance handle."
         tgt_inst.display_logo_error_msg()
      Case 32764
          tgt_inst.logo_error_text = "32764 - The ENABLETEMPLATE flag was set in the Flags member of a common dialog data structure " & "but the application failed to provide a corresponding template."
         tgt_inst.display_logo_error_msg()
      Case 32765
          tgt_inst.logo_error_text = "32765 - The common dialog function failed during initialization."
         tgt_inst.display_logo_error_msg()
      Case 32766
          tgt_inst.logo_error_text = "32766 - The IStructSize member of the corresponding common dialog data structure is invalid."
         tgt_inst.display_logo_error_msg()
      Case 35600
          tgt_inst.logo_error_text = "35600 - Index out of bounds"
         tgt_inst.display_logo_error_msg()
      Case 35601
          tgt_inst.logo_error_text = "35601 - Element not found"
         tgt_inst.display_logo_error_msg()
      Case 35602
          tgt_inst.logo_error_text = "35602 - Key is not unique in Collection"
         tgt_inst.display_logo_error_msg()
      Case 35603
          tgt_inst.logo_error_text = "35603 - Invalid Key"
         tgt_inst.display_logo_error_msg()
      Case 35604
          tgt_inst.logo_error_text = "35604 - When the Listview Control's View property is set to 3 (report), " & " the left-most column (1) can only be left aligned. Any attempt to set the alignment " & " to another value will result in this error."
         tgt_inst.display_logo_error_msg()
      Case 35605
          tgt_inst.logo_error_text = "35605 - This item's control has been deleted."
         tgt_inst.display_logo_error_msg()
      Case 35606
          tgt_inst.logo_error_text = "35606 - Control's Collection has been modified."
         tgt_inst.display_logo_error_msg()
      Case 35607
          tgt_inst.logo_error_text = "35607 - Required Argument is missing."
         tgt_inst.display_logo_error_msg()
      Case 35610
          tgt_inst.logo_error_text = "35610 - Invalid Object"
         tgt_inst.display_logo_error_msg()
      Case 35611
          tgt_inst.logo_error_text = "35611 - Property is read-only if image list contains Images."
         tgt_inst.display_logo_error_msg()
      Case 35613
          tgt_inst.logo_error_text = "35613 - ImageList must be initialized before it can be used."
         tgt_inst.display_logo_error_msg()
      Case 35614
          tgt_inst.logo_error_text = "35614 - This would introduce a cycle."
         tgt_inst.display_logo_error_msg()
      Case 35615
          tgt_inst.logo_error_text = "35615 - All images in list must be same size."
         tgt_inst.display_logo_error_msg()
      Case 35616
          tgt_inst.logo_error_text = "35616 - Maximum panels exceeded"
         tgt_inst.display_logo_error_msg()
      Case 35617
          tgt_inst.logo_error_text = "35617 - ImageList cannot be modified while another control is bound to it."
         tgt_inst.display_logo_error_msg()
      Case 35619
          tgt_inst.logo_error_text = "35619 - Maximum buttons exceeded"
         tgt_inst.display_logo_error_msg()
      Case 35700
          tgt_inst.logo_error_text = "35700 - Circular object referencing is not allowed"
         tgt_inst.display_logo_error_msg()
          'over 39000 values'
      Case 50030
          tgt_inst.logo_error_text = "50030 - Unexpected Error"
         tgt_inst.display_logo_error_msg()
  End Select

End Sub

End Module