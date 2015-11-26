VERSION 5.00
Begin VB.Form frmSearchResults 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Please choose a script locations"
   ClientHeight    =   3780
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   1215
   End
   Begin VB.ListBox List_Locations 
      Height          =   3765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmSearchResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public SelectedLocation&
Public Locations As Collection


Private Sub cmd_cancel_Click()
   DoCancel
End Sub

Private Sub cmd_ok_Click()
   DoSelect
End Sub

Private Sub Form_Initialize()
  
   SelectedLocation = -1

End Sub

Private Sub Form_Load()
   
   With List_Locations
      Dim item
      For Each item In Locations
         .AddItem H32(item)
      Next
      
      If .ListCount >= 1 Then
         .ListIndex = 0
      End If
      
   End With

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   FrmMain.WH_close
   If UnloadMode = 0 Then DoCancel
End Sub


Private Sub List_Locations_DblClick()
   DoSelect
End Sub

Private Sub List_Locations_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   FrmMain.WH_Open
   FrmMain.WH_Goto HexToInt(List_Locations.Text)
End Sub

Public Sub DoSelect()
   SelectedLocation = List_Locations.ListIndex + 1
   Unload Me
End Sub

Public Sub DoCancel()
   SelectedLocation = -1
   Unload Me
End Sub
