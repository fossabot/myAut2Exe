VERSION 5.00
Begin VB.Form FrmFuncRename_StringMismatch 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Function Renamer String Mismatch"
   ClientHeight    =   4035
   ClientLeft      =   2445
   ClientTop       =   2430
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      ToolTipText     =   "Press 'esc' to reject this pair of functions"
      Top             =   0
      Width           =   6015
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Press 'enter' to accept this"
      Top             =   0
      Width           =   5895
   End
   Begin VB.ListBox List_Inc 
      Appearance      =   0  'Flat
      Height          =   8610
      ItemData        =   "FrmFuncRename_StringMismatch.frx":0000
      Left            =   6120
      List            =   "FrmFuncRename_StringMismatch.frx":0002
      TabIndex        =   1
      Top             =   360
      Width           =   5895
   End
   Begin VB.ListBox List_Org 
      Appearance      =   0  'Flat
      Height          =   8610
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "FrmFuncRename_StringMismatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum AcceptResult_enum
   Result_True
   Result_False
   Result_Undefined
End Enum

Private mAcceptResult As AcceptResult_enum
Public Property Get AcceptResult() As AcceptResult_enum
   AcceptResult = mAcceptResult
End Property



Public Sub Create(fn_org As MatchCollection, fn_inc As MatchCollection)
   FillList List_Org, fn_org
   FillList List_Inc, fn_inc
   
   mAcceptResult = Result_Undefined
End Sub

Private Sub FillList(List As Listbox, Match As MatchCollection)
   List.Clear
   
   Dim i As Match
   For Each i In Match
      List.AddItem i '.SubMatches(1)
   Next
End Sub

Private Sub cmd_cancel_Click()
   Unload Me
   mAcceptResult = Result_False
End Sub

Private Sub cmd_ok_Click()
   Unload Me
   mAcceptResult = Result_True
End Sub

