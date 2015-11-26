VERSION 5.00
Begin VB.Form FrmAHK_KeyFinder 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "AHK Keyfinder"
   ClientHeight    =   1500
   ClientLeft      =   6300
   ClientTop       =   4770
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Cmd_ok 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox Combo_AHK_Key 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Use Mousewheel or the up and down keys to scroll. Double click to restore inital key."
      Height          =   855
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   1695
   End
End
Attribute VB_Name = "FrmAHK_KeyFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ScriptDataEncryptedPreviewBuff As StringReader
Dim AHK_Key_initial As Byte

Public AHK_Key As Byte

Public Sub Create(ScriptData As StringReader, AHK_Key As Byte)

   Set ScriptDataEncryptedPreviewBuff = New StringReader
   
   ScriptData.Position = 0
   ScriptDataEncryptedPreviewBuff.Data = ScriptData.FixedString(2000)
   
   With Combo_AHK_Key
      Dim i&
      For i = 0 To &HFF
         .AddItem H8(i)
      Next
      
   End With

   AHK_Key_initial = AHK_Key
   Form_DblClick

End Sub

Private Sub cmd_cancel_Click()
   AHK_Key = AHK_Key_initial
   Unload Me
End Sub

Private Sub cmd_ok_Click()
   Unload Me
End Sub

Private Sub Combo_AHK_Key_Change()
   Combo_AHK_Key_Click
End Sub

Private Sub Combo_AHK_Key_Click()
   On Error GoTo Combo_AHK_Key_Change_err
   
   AHK_Key = HexToInt(Combo_AHK_Key.Text)
'   FrmMain.Txt_Script =
AHK_ExtraDecryption ScriptDataEncryptedPreviewBuff, AHK_Key

Combo_AHK_Key_Change_err:
End Sub


Private Sub Form_DblClick()
    Combo_AHK_Key.ListIndex = AHK_Key_initial
End Sub

'Private Sub Combo_AHK_Key_KeyUp(KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
'      Unload Me
'   End If
'End Sub



'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'   If UnloadMode <> QueryUnloadConstants.vbFormCode Then
'      cmd_cancel_Click
'   End If
'End Sub

