VERSION 5.00
Begin VB.Form Frm_Options 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8400
   LinkTopic       =   "Options"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Get XorKey's"
      Height          =   855
      Left            =   120
      TabIndex        =   54
      Top             =   6120
      Width           =   8175
      Begin VB.TextBox Txt_GetCamoFileName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   56
         ToolTipText     =   "FileName that is used when you click GetCamo's"
         Top             =   450
         Width           =   6255
      End
      Begin VB.CommandButton cmd_CamoGet 
         Caption         =   "GetCamo's"
         Height          =   375
         Left            =   120
         TabIndex        =   55
         ToolTipText     =   "Click this if Au3-Stub was modifies by AutoIt3Camo - Requires UNPACKED(i.g. No UPX) Data to work!"
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmd_ResetOptions 
      Caption         =   "Reset Options"
      Height          =   375
      Left            =   2040
      TabIndex        =   53
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmd_ExportSettings 
      Caption         =   "Export settings to file"
      Height          =   375
      Left            =   120
      TabIndex        =   50
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox txt_AU2_Type 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   6240
      TabIndex        =   45
      Tag             =   "AU2!"
      Text            =   "AU2!"
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_AU2_Type_Hex 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      TabIndex        =   44
      Text            =   "11 22 33 44"
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "ScriptBody XORKey's"
      Height          =   2535
      Left            =   120
      TabIndex        =   29
      Top             =   2280
      Width           =   8175
      Begin VB.TextBox txt_AU3_ResTypeFile_hex 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtCompiledPathName_DataNew 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   18
         Tag             =   "0F479"
         Text            =   "0F479"
         ToolTipText     =   "0F479"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtCompiledPathName_LenNew 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Tag             =   "0F820"
         Text            =   "0F820"
         ToolTipText     =   "0F820 Path and Filename that was used to compile the script/Fileinstall resource"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtCompiledPathName_Data 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         TabIndex        =   25
         Tag             =   "F25E"
         Text            =   "F25E"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtCompiledPathName_Len 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         TabIndex        =   24
         Tag             =   "29AC"
         Text            =   "29AC"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtSrcFile_FileInst_DataNew 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   16
         Tag             =   "B33F"
         Text            =   "B33F"
         ToolTipText     =   "B33F"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtSrcFile_FileInst_LenNew 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Tag             =   "ADBC"
         Text            =   "ADBC"
         ToolTipText     =   "ADBC The SrcFile_FileInst is normally '>>>AUTOIT SCRIPT<<<' "
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtSrcFile_FileInst_Data 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         TabIndex        =   23
         Tag             =   "A25E"
         Text            =   "A25E"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtSrcFile_FileInst_Len 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         TabIndex        =   22
         Tag             =   "29BC"
         Text            =   "29BC"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtData_DecryptionKey_New 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         TabIndex        =   14
         Tag             =   "2477"
         Text            =   "2477"
         ToolTipText     =   "2477 DecryptionKey"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtData_DecryptionKey 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         TabIndex        =   21
         Tag             =   "22AF"
         Text            =   "22AF"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtXORKey_MD5PassphraseHashText_DataNew 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   13
         Tag             =   "99F2"
         Text            =   "99F2"
         ToolTipText     =   "99F2"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtXORKey_MD5PassphraseHashText_Data 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         TabIndex        =   20
         Tag             =   "C3D2"
         Text            =   "C3D2"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtXORKey_MD5PassphraseHashText_Len 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         TabIndex        =   19
         Tag             =   "FAC1"
         Text            =   "FAC1"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txt_AU3_ResTypeFile 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Tag             =   "FILE"
         Text            =   "FILE"
         ToolTipText     =   "The 'FILE' marker marks the beginning of the mainscript and every File that is installed by FILEINSTALL"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txt_FILE_DecryptionKey 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   12
         Tag             =   "18EE"
         Text            =   "18EE"
         ToolTipText     =   "18EE FILE-decryptionKey - normally there should be no reason to touch this."
         Top             =   360
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   5040
         X2              =   5040
         Y1              =   720
         Y2              =   2280
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "AddKey"
         Height          =   255
         Left            =   7200
         TabIndex        =   48
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "AddKey"
         Height          =   255
         Left            =   3960
         TabIndex        =   47
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "CompiledPathName"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "SrcFile_FileInst"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         Height          =   255
         Left            =   3000
         TabIndex        =   41
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Length"
         Height          =   255
         Left            =   2040
         TabIndex        =   40
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         Height          =   255
         Left            =   6240
         TabIndex        =   39
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Length"
         Height          =   255
         Left            =   5280
         TabIndex        =   38
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "MD5PassphraseHashText"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...for AutoIt 3.2.6 and newer"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   36
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "... for older AU3, AHK and AU2 "
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   35
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "AU3 ResourceTypeFILE "
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Other Options"
      Height          =   615
      Left            =   120
      TabIndex        =   27
      Top             =   4920
      Width           =   8175
      Begin VB.CheckBox chk_extractIcon 
         Caption         =   "Extract Icon"
         Height          =   255
         Left            =   3480
         TabIndex        =   57
         ToolTipText     =   "Deselect if you don't need the icon(*.ico) file; Grey - don't keep Au3-Standard *.ico"
         Top             =   240
         UseMaskColor    =   -1  'True
         Value           =   2  'Grayed
         Width           =   1215
      End
      Begin VB.CheckBox chk_disableWinhex 
         Caption         =   "Disable Winhex"
         Height          =   255
         Left            =   1920
         TabIndex        =   51
         ToolTipText     =   "Disables Winhex since it maybe disturbing that eac time you change the start offset it'll pops up."
         Top             =   240
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox Chk_RestoreIncludes 
         Caption         =   "Restore Includes"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         ToolTipText     =   "That will only work for AHK and old not tokenise AU3"
         Top             =   240
         Value           =   1  'Checked
         Width           =   1560
      End
   End
   Begin VB.Frame Fr_ScriptStart 
      Caption         =   "ScriptStart"
      Height          =   1935
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   8175
      Begin VB.TextBox txt_AU3Sig_Hex 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   52
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox txt_AU3_Type_hex 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Txt_AU2_SubType_hex 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5400
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txt_AU2_SubType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   8
         Tag             =   "EA05"
         Text            =   "EA05"
         ToolTipText     =   "Comes right after AU3-Type and is used a second time as signature for compressed data"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox Txt_AU3_SubType_hex 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txt_AU3_SubType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Tag             =   "EA06"
         Text            =   "EA06"
         ToolTipText     =   "Comes right after AU3-Type and is used a second time as signature for compressed data"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txt_AU3_Type 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Tag             =   "AU3!"
         Text            =   "AU3!"
         ToolTipText     =   "This follows just after the Au3_Signature "
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txt_AU3Sig 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Tag             =   $"Frm_Options.frx":0000
         Text            =   $"Frm_Options.frx":0014
         ToolTipText     =   "Only imporatant if 'Use 'normal' Au3_Signature to find start of script' is checked"
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox Chk_ForceOldScriptType 
         Caption         =   "Force Old Script Type"
         Height          =   255
         Left            =   4200
         TabIndex        =   2
         ToolTipText     =   "Grey means auto detect and is the best in most cases. Background: AHK and AU2 don't have the 'SrcFile_FileInst' in the Header."
         Top             =   240
         Value           =   2  'Grayed
         Width           =   2295
      End
      Begin VB.CheckBox Chk_NormalSigScan 
         Caption         =   "Use 'normal' Au3_Signature to find start of script"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "AU2 SubType"
         Height          =   255
         Left            =   3600
         TabIndex        =   33
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "AU3 SubType"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "AU3 Type"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "AU3 Signature"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CheckBox Chk_NoDeTokenise 
      Caption         =   "Disable Detokeniser"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Enable that when you decompile AutoItScripts lower than ver 3.1.6"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label18 
      Caption         =   "Note: Settings are saved to registry when you close this window. "
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   6000
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "AU2_Type"
      Height          =   255
      Left            =   5040
      TabIndex        =   46
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Frm_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bDisableTextChangeEvent As Boolean
Private Const ValidTypeStringLen = 4
Private Const ValidSigStringLen = 16

Private Const ValidHexStringLen = 8

Const Setting_ExcludeFromLoadSave$ = _
   "Txt_GetCamoFileName" & " " & _
   "" & _
   "txt_AU3Sig" & " " & _
   "txt_AU3_Type" & " " & _
   "txt_AU3_SubType" & " " & _
   "txt_AU2_SubType" & " " & _
   "txt_AU3_ResTypeFile" & " " & _
   "" & _
   "txt_AU2_Type" & " " & _
   "txt_AU2_Type_Hex" & " " & _
   "Chk_NoDeTokenise" & " " & _
   ""

Private Sub RegRun(verb, ExportFileName, Optional options = "")
   
   Dim Regpath$
   Regpath = Quote("HKCU\Software\VB and VBA Program Settings\" & App.Title & "\" & Me.Name)
   
   Const cmd$ = "Reg.exe"
   
   Log "Running: " & Quote(cmd & " " & _
                           verb & " " & _
                           Regpath & " " & _
                           ExportFileName & " " & _
                           options)
   Dim retval&
   retval = ShellEx(cmd, verb & " " & _
                         Regpath & " " & _
                         ExportFileName & " " & _
                         options, vbHide) 'vbNormalFocus )

End Sub


Private Sub chk_extractIcon_Click()
   Static value
   Checkbox_TriStateToggle chk_extractIcon, value
End Sub

Private Sub cmd_CamoGet_Click()
   On Error GoTo cmd_CamoGet_Click_err
   
   CamoGet
   
   Exit Sub
cmd_CamoGet_Click_err:
   Log "ERROR: [CamoGet] " & Err.Description
End Sub

Private Sub cmd_ExportSettings_Click()
   Dim ExportFileName As New ClsFilename
   
   On Error Resume Next

 ' Save Setting to registry
   Settings_save
   
   With ExportFileName
      .NameWithExt = "myAut2Exe-Settings"
      If FileExists(FileName) Then
         .NameWithExt = FileName.Name & "_" & .NameWithExt
      End If
      
      .NameWithExt = InputBox("ExportFileName:", _
                     "Save Settings to File", _
                     .NameWithExt)
      .Ext = "reg"
   End With
   
   RegRun "EXPORT", ExportFileName
                         
   Log Quote(CurDir & "\" & ExportFileName) & " created."

End Sub


Private Sub cmd_ResetOptions_Click()
   On Error Resume Next

 ' Delete Registry settings
   RegRun "DELETE", "", "/va /f"


 ' Restore Defaults
'   Form_Load
 
   Unload Me
   
   Me.Show

End Sub




Private Sub Form_Activate()
   Txt_GetCamoFileName = FrmMain.Combo_Filename
End Sub

Private Sub Form_Load()

   FormSettings_Load Me, Setting_ExcludeFromLoadSave
   
   LocaleID = ConfigValue_Load(Me.Name, "LocaleID", LocaleID)
   
   
 ' Ensure Initialisation
   CommitChanges


End Sub

'!!! Important !!!
' Call CommitChanges every time you made changes to the form like for ex. this:
' Frm_Options.txt_AU3Sig.Text = "AU3"
Public Sub CommitChanges()

   Dim dummy As Boolean
   
 ' Script Start
   txt_AU3Sig_Hex_Validate dummy
   txt_AU3_Type_hex_Validate dummy
'   txt_AU2_Type_hex_Validate dummy
   txt_AU3_SubType_hex_Validate dummy
   txt_AU2_SubType_hex_Validate dummy
   
  
 ' ScriptBody XORKey's
   
   txt_AU3_ResTypeFile_hex_Validate dummy
   txt_FILE_DecryptionKey_Validate dummy

   txtXORKey_MD5PassphraseHashText_Len_Validate dummy
   txtXORKey_MD5PassphraseHashText_Data_Validate dummy
   txtData_DecryptionKey_Validate dummy
   
   txtXORKey_MD5PassphraseHashText_DataNEW_Validate dummy
   txtData_DecryptionKey_New_Validate dummy
   
   
   txtSrcFile_FileInst_Len_Validate dummy
   txtSrcFile_FileInst_Data_Validate dummy
   
   txtSrcFile_FileInst_LenNew_Validate dummy
   txtSrcFile_FileInst_DataNew_Validate dummy
   
   txtCompiledPathName_Len_Validate dummy
   txtCompiledPathName_Data_Validate dummy
   
   txtCompiledPathName_LenNew_Validate dummy
   txtCompiledPathName_DataNew_Validate dummy



End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode <> vbFormCode Then
      Settings_save
   End If
End Sub
   
   
Private Sub Settings_save()
   ConfigValue_Save(Me.Name, "LocaleID") = LocaleID
   FormSettings_Save Me, Setting_ExcludeFromLoadSave
End Sub



''////////////////////////////////////////////////////////
''// Text AU2_Type
'Private Sub txt_AU2_Type_Change()
'   txt_Changed txt_AU2_Type, txt_AU2_Type_Hex, ValidTypeStringLen
'End Sub
'Private Sub txt_AU2_Type_Validate(Cancel As Boolean)
'   txt_Validate txt_AU2_Type, AU2_TypeStr
'End Sub
'
'Private Sub txt_AU2_Type_hex_Change()
'   txtHex_Changed txt_AU2_Type, txt_AU2_Type_Hex, ValidTypeStringLen
'End Sub
'Private Sub txt_AU2_Type_hex_Validate(Cancel As Boolean)
'   txtHex_Validate txt_AU2_Type, txt_AU2_Type_Hex, AU2_TypeStr, ValidTypeStringLen
'End Sub

'////////////////////////////////////////////////////////
'// Text AU2_SubType
Private Sub txt_AU2_SubType_Change()
   txt_Changed txt_AU2_SubType, Txt_AU2_SubType_hex, ValidTypeStringLen
End Sub
Private Sub txt_AU2_SubType_Validate(Cancel As Boolean)
   txt_Validate txt_AU2_SubType, AU3_SubTypeStr_old
End Sub

Private Sub txt_AU2_SubType_hex_Change()
   txtHex_Changed txt_AU2_SubType, Txt_AU2_SubType_hex, ValidTypeStringLen
End Sub
Private Sub txt_AU2_SubType_hex_Validate(Cancel As Boolean)
   txtHex_Validate txt_AU2_SubType, Txt_AU2_SubType_hex, AU3_SubTypeStr_old, ValidTypeStringLen
End Sub


'////////////////////////////////////////////////////////
'// Text AU3_SubType
Private Sub txt_AU3_SubType_Change()
   txt_Changed txt_AU3_SubType, Txt_AU3_SubType_hex, ValidTypeStringLen
End Sub
Private Sub txt_AU3_SubType_Validate(Cancel As Boolean)
   txt_Validate txt_AU3_SubType, AU3_SubTypeStr
End Sub

Private Sub txt_AU3_SubType_hex_Change()
   txtHex_Changed txt_AU3_SubType, Txt_AU3_SubType_hex, ValidTypeStringLen
End Sub
Private Sub txt_AU3_SubType_hex_Validate(Cancel As Boolean)
   txtHex_Validate txt_AU3_SubType, Txt_AU3_SubType_hex, AU3_SubTypeStr, ValidTypeStringLen
End Sub


'////////////////////////////////////////////////////////
'// Text AU3_Type
Private Sub txt_AU3_Type_Change()
   txt_Changed txt_AU3_Type, txt_AU3_Type_hex, ValidTypeStringLen
End Sub
Private Sub txt_AU3_Type_Validate(Cancel As Boolean)
   txt_Validate txt_AU3_Type, AU3_TypeStr
End Sub

Private Sub txt_AU3_Type_hex_Change()
   txtHex_Changed txt_AU3_Type, txt_AU3_Type_hex, ValidTypeStringLen
End Sub
Private Sub txt_AU3_Type_hex_Validate(Cancel As Boolean)
   txtHex_Validate txt_AU3_Type, txt_AU3_Type_hex, AU3_TypeStr, ValidTypeStringLen
End Sub



'////////////////////////////////////////////////////////
'// Text AU3Sig
Private Sub txt_AU3Sig_Change()
   txt_Changed txt_AU3Sig, txt_AU3Sig_Hex, ValidSigStringLen
End Sub
Private Sub txt_AU3Sig_Validate(Cancel As Boolean)
   txt_Validate txt_AU3Sig, AU3Sig_HexStr
   AU3Sig_HexStr = txt_AU3Sig_Hex.Text
End Sub

Private Sub txt_AU3Sig_Hex_Change()
   txtHex_Changed txt_AU3Sig, txt_AU3Sig_Hex, ValidSigStringLen
End Sub
Private Sub txt_AU3Sig_Hex_Validate(Cancel As Boolean)
   txtHex_Validate txt_AU3Sig, txt_AU3Sig_Hex, AU3Sig_HexStr, ValidSigStringLen
   AU3Sig_HexStr = txt_AU3Sig_Hex.Text
End Sub


'////////////////////////////////////////////////////////
'// Text AU3_ResTypeFile
Private Sub txt_AU3_ResTypeFile_Change()
   txt_Changed txt_AU3_ResTypeFile, txt_AU3_ResTypeFile_hex, ValidTypeStringLen
End Sub
Private Sub txt_AU3_ResTypeFile_Validate(Cancel As Boolean)
   txt_Validate txt_AU3_ResTypeFile, AU3_ResTypeFile
End Sub

Private Sub txt_AU3_ResTypeFile_hex_Change()
   txtHex_Changed txt_AU3_ResTypeFile, txt_AU3_ResTypeFile_hex, ValidTypeStringLen
End Sub
Private Sub txt_AU3_ResTypeFile_hex_Validate(Cancel As Boolean)
   txtHex_Validate txt_AU3_ResTypeFile, txt_AU3_ResTypeFile_hex, AU3_ResTypeFile, ValidTypeStringLen
End Sub



'///////////////////////////////


Private Sub txt_Changed(TextBox As TextBox, TextBoxHex As TextBox, ValidDataLen&)
   On Error GoTo txt_Changed_err
   
   If bDisableTextChangeEvent Then Exit Sub
   bDisableTextChangeEvent = True
 
 ' Set Hexvalues
   Dim myString As New StringReader
   myString = TextBox
 
 ' TextBoxHex.Tag is there to hold strings containing 00-bytes ;since you can not enter 00 into a Textbox
 ' When text is inputed TextBoxHex.Tag in cleared
   TextBoxHex.Tag = ""
   TextBoxHex = Trim(ValuesToHexString(myString))
 ' Mark Hexvalues as valid
   TxtSetValidAndDefault(TextBoxHex, True) = True

txt_Changed_err:
   TxtSetValidAndDefault(TextBox, txt_IsDefault(TextBox)) = txt_AUX_Type_IsValid(TextBox.Text, ValidDataLen)
   bDisableTextChangeEvent = False
End Sub


Private Sub txt_Validate(TextBox As TextBox, ByRef OutputVal$)
   
   With TextBox
    
    ' restore default if invalid
      If TxtSetValidAndDefault(TextBox, False) = False Then
         .Text = .Tag
      End If
         
      OutputVal = .Text

   End With
End Sub


Private Sub txtHex_Changed(TextBox As TextBox, TextBoxHex As TextBox, ValidDataLen&)
   
   On Error GoTo txtHex_Changed_err
   
   If bDisableTextChangeEvent Then Exit Sub
   bDisableTextChangeEvent = True
   
   ' Hexvalues -> String
     TextBoxHex.Tag = HexvaluesToString(TextBoxHex.Text)
     TextBox = TextBoxHex.Tag
   
txtHex_Changed_err:
    ' Possible Hex input errors?
      TxtSetValidAndDefault(TextBoxHex, True) = (Err = 0)
    
    ' Validate Text?
      TxtSetValidAndDefault(TextBox, txt_IsDefault(TextBox)) = txt_AUX_Type_IsValid(TextBoxHex.Tag, ValidDataLen)
      
   
   bDisableTextChangeEvent = False

End Sub

Private Sub txtRestoreDefault(TextBox As TextBox)
   With TextBox
    ' Note: 'TextBox.Text = ' will trigger the Text_Change event
       ' ...but not when they are equal(due to stupid VB-Design)
       If .Text = .Tag Then .Text = ""
      .Text = .Tag
      .Text = TextBox.Tag
   End With
End Sub
Private Sub txtHex_Validate(TextBox As TextBox, TextBoxHex As TextBox, ByRef OutputVal$, ValidDataLen&)
   On Error GoTo txtHex_Validate_err
 
 ' Call changed for first time init
   txtHex_Changed TextBox, TextBoxHex, ValidDataLen
  
 ' Restore default value if invalid
   If TxtSetValidAndDefault(TextBox, False) = False Then
      
      txtRestoreDefault TextBox
      OutputVal = TextBox.Tag
      
      Exit Sub
   Else
   
    ' Block 'txtHex_Validate' without anything previous 'txtHex_Change'
      If TextBoxHex.Tag = "" Then Exit Sub

      OutputVal = TextBoxHex.Tag
   
   End If
 
 ' Set Hexvalues
   Dim myString As New StringReader
   myString = TextBoxHex.Tag
   TextBoxHex = Trim(ValuesToHexString(myString))
 
 ' Mark Hexvalues as valid
   TxtSetValidAndDefault(TextBoxHex, True) = True

txtHex_Validate_err:
   TxtSetValidAndDefault(TextBox, txt_IsDefault(TextBox)) = _
      txt_AUX_Type_IsValid(TextBoxHex.Tag, ValidDataLen)
End Sub

'////////////////////////////
'// TxtH32SetValidAndDefault
Property Get TxtSetValidAndDefault(TextBox As TextBox, bIsDefault As Boolean) As Boolean
   TxtSetValidAndDefault = TextBox.ForeColor <> vbRed
End Property
Property Let TxtSetValidAndDefault(TextBox As TextBox, bIsDefault As Boolean, bIsValid As Boolean)
   With TextBox
   
   ' bState = false  ->   Red   => invalid
   ' bState = True   ->   Black =>   valid Default
   ' bState = True   ->   Blue  =>   valid nonDefault
      .ForeColor = IIf(bIsValid, _
                          IIf(bIsDefault, vbBlack, vbBlue), _
                          vbRed)
   End With
End Property



Function txt_H32_IsValid(TextBox As TextBox) As Boolean
   With TextBox
      On Error Resume Next
      H32 HexToInt(.Text)
      txt_H32_IsValid = (Err = 0) And _
                        (Len(.Text) <= ValidHexStringLen)
   End With
End Function
Function txt_AUX_Type_IsValid(checkText$, ValidDataLen&) As Boolean
   txt_AUX_Type_IsValid = (Len(checkText) = ValidDataLen)
End Function


Function txt_H32_IsDefault(TextBox As TextBox) As Boolean
   On Error Resume Next
   With TextBox
      txt_H32_IsDefault = (HexToInt(.Text) = HexToInt(.Tag))
   End With
End Function

Function txt_IsDefault(TextBox As TextBox) As Boolean
   With TextBox
      txt_IsDefault = (.Text = .Tag)
   End With
End Function



Private Sub txt_H32_Validate(TextBox As TextBox, ByRef OutputVal&)
   With TextBox
      
      If txt_H32_IsValid(TextBox) Then
         .Text = H32(HexToInt(.Text))
        
       ' If default - set default value(sounds stupid but is good to set default with format like no leading 0's)
         If txt_H32_IsDefault(TextBox) Then .Text = .Tag 'txtRestoreDefault TextBox
    
      Else
         'Invalid -> restore default
         .Text = .Tag 'txtRestoreDefault TextBox
         
      End If
      
      OutputVal = HexToInt(.Text)

   End With
End Sub









Private Sub Chk_ForceOldScriptType_Click()
   Static value
   Checkbox_TriStateToggle Chk_ForceOldScriptType, value
End Sub

'////////////////////////////////
'// FILE_DecryptionKey
Private Sub txt_FILE_DecryptionKey_Change()
   Dim t As TextBox
   Set t = txt_FILE_DecryptionKey
   
   TxtSetValidAndDefault(t, txt_H32_IsDefault(t)) = txt_H32_IsValid(t)

End Sub

Private Sub txt_FILE_DecryptionKey_Validate(Cancel As Boolean)
   txt_H32_Validate txt_FILE_DecryptionKey, FILE_DecryptionKey
End Sub











'///////////////////////////////////////////
'// txtXORKey_MD5PassphraseHashText_Len
Private Sub txtXORKey_MD5PassphraseHashText_Len_Change()
   Dim t As TextBox
   Set t = txtXORKey_MD5PassphraseHashText_Len
   
   TxtSetValidAndDefault(t, txt_H32_IsDefault(t)) = txt_H32_IsValid(t)
End Sub

Private Sub txtXORKey_MD5PassphraseHashText_Len_Validate(Cancel As Boolean)
   txt_H32_Validate txtXORKey_MD5PassphraseHashText_Len, XORKey_MD5PassphraseHashText_Len
End Sub


'///////////////////////////////////////////
'// txtXORKey_MD5PassphraseHashText_Data
Private Sub txtXORKey_MD5PassphraseHashText_Data_Change()
   Dim t As TextBox
   Set t = txtXORKey_MD5PassphraseHashText_Data
   
   TxtSetValidAndDefault(t, txt_H32_IsDefault(t)) = txt_H32_IsValid(t)

End Sub

Private Sub txtXORKey_MD5PassphraseHashText_Data_Validate(Cancel As Boolean)
   txt_H32_Validate txtXORKey_MD5PassphraseHashText_Data, XORKey_MD5PassphraseHashText_Data
End Sub

'///////////////////////////////////////////
'// txtData_DecryptionKey
Private Sub txtData_DecryptionKey_Change()
   Dim t As TextBox
   Set t = txtData_DecryptionKey
   
   TxtSetValidAndDefault(t, txt_H32_IsDefault(t)) = txt_H32_IsValid(t)

End Sub

Private Sub txtData_DecryptionKey_Validate(Cancel As Boolean)
   txt_H32_Validate txtData_DecryptionKey, Data_DecryptionKey
End Sub




'///////////////////////////////////////////
'// txtXORKey_MD5PassphraseHashText_DataNEW
Private Sub txtXORKey_MD5PassphraseHashText_DataNEW_Change()
   Dim t As TextBox
   Set t = txtXORKey_MD5PassphraseHashText_DataNew
   
   TxtSetValidAndDefault(t, txt_H32_IsDefault(t)) = txt_H32_IsValid(t)

End Sub

Private Sub txtXORKey_MD5PassphraseHashText_DataNEW_Validate(Cancel As Boolean)
   txt_H32_Validate txtXORKey_MD5PassphraseHashText_DataNew, XORKey_MD5PassphraseHashText_DataNEW
End Sub
'///////////////////////////////////////////
'// txtData_DecryptionKey
Private Sub txtData_DecryptionKey_New_Change()
   Dim t As TextBox
   Set t = txtData_DecryptionKey_New
   
   TxtSetValidAndDefault(t, txt_H32_IsDefault(t)) = txt_H32_IsValid(t)

End Sub

Private Sub txtData_DecryptionKey_New_Validate(Cancel As Boolean)
   txt_H32_Validate txtData_DecryptionKey_New, Data_DecryptionKey_NewConst
End Sub



'///////////////////////////////////////////
'// txtSrcFile_FileInst
Private Sub txtSrcFile_FileInst_Len_Change()
   Dim t As TextBox
   Set t = txtSrcFile_FileInst_Len
   
   TxtSetValidAndDefault(t, txt_H32_IsDefault(t)) = txt_H32_IsValid(t)

End Sub
Private Sub txtSrcFile_FileInst_Len_Validate(Cancel As Boolean)
   txt_H32_Validate txtSrcFile_FileInst_Len, Xorkey_SrcFile_FileInst_Len
End Sub

Private Sub txtSrcFile_FileInst_Data_Change()
   Dim t As TextBox
   Set t = txtSrcFile_FileInst_Data
   
   TxtSetValidAndDefault(t, txt_H32_IsDefault(t)) = txt_H32_IsValid(t)

End Sub

Private Sub txtSrcFile_FileInst_Data_Validate(Cancel As Boolean)
   txt_H32_Validate txtSrcFile_FileInst_Data, Xorkey_SrcFile_FileInst_Data
End Sub

'///////////////////////////////////////////
'// txtSrcFile_FileInstNEW
Private Sub txtSrcFile_FileInst_LenNew_Change()
   Dim t As TextBox
   Set t = txtSrcFile_FileInst_LenNew
   
   TxtSetValidAndDefault(t, txt_H32_IsDefault(t)) = txt_H32_IsValid(t)

End Sub
Private Sub txtSrcFile_FileInst_LenNew_Validate(Cancel As Boolean)
   txt_H32_Validate txtSrcFile_FileInst_LenNew, Xorkey_SrcFile_FileInstNEW_Len
End Sub

Private Sub txtSrcFile_FileInst_DataNew_Change()
   Dim t As TextBox
   Set t = txtSrcFile_FileInst_DataNew
   
   TxtSetValidAndDefault(t, txt_H32_IsDefault(t)) = txt_H32_IsValid(t)

End Sub

Private Sub txtSrcFile_FileInst_DataNew_Validate(Cancel As Boolean)
   txt_H32_Validate txtSrcFile_FileInst_DataNew, Xorkey_SrcFile_FileInstNEW_Data
End Sub


'///////////////////////////////////////////
'// txtCompiledPathName
Private Sub txtCompiledPathName_Len_Change()
   Dim t As TextBox
   Set t = txtCompiledPathName_Len
   
   TxtSetValidAndDefault(t, txt_H32_IsDefault(t)) = txt_H32_IsValid(t)

End Sub
Private Sub txtCompiledPathName_Len_Validate(Cancel As Boolean)
   txt_H32_Validate txtCompiledPathName_Len, Xorkey_CompiledPathName_Len
End Sub

Private Sub txtCompiledPathName_Data_Change()
   Dim t As TextBox
   Set t = txtCompiledPathName_Data
   
   TxtSetValidAndDefault(t, txt_H32_IsDefault(t)) = txt_H32_IsValid(t)

End Sub

Private Sub txtCompiledPathName_Data_Validate(Cancel As Boolean)
   txt_H32_Validate txtCompiledPathName_Data, Xorkey_CompiledPathName_Data
End Sub



'///////////////////////////////////////////
'// txtCompiledPathNameNEW
Private Sub txtCompiledPathName_LenNew_Change()
   Dim t As TextBox
   Set t = txtCompiledPathName_LenNew
   
   TxtSetValidAndDefault(t, txt_H32_IsDefault(t)) = txt_H32_IsValid(t)

End Sub
Private Sub txtCompiledPathName_LenNew_Validate(Cancel As Boolean)
   txt_H32_Validate txtCompiledPathName_LenNew, Xorkey_CompiledPathNameNEW_Len
End Sub

Private Sub txtCompiledPathName_DataNew_Change()
   Dim t As TextBox
   Set t = txtCompiledPathName_DataNew
   
   TxtSetValidAndDefault(t, txt_H32_IsDefault(t)) = txt_H32_IsValid(t)

End Sub

Private Sub txtCompiledPathName_DataNew_Validate(Cancel As Boolean)
   txt_H32_Validate txtCompiledPathName_DataNew, Xorkey_CompiledPathNameNEW_Data
End Sub



