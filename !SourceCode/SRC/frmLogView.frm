VERSION 5.00
Begin VB.Form frmLogView 
   Caption         =   "View Log"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11670
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtlog 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.CommandButton cmd_Quit 
      Cancel          =   -1  'True
      Caption         =   "Quit"
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmLogView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Quit_Click()
   Me.Hide
End Sub

Private Sub Form_Resize()
On Error Resume Next
   
   With txtlog
   .Height = Me.Height - 550
   .Width = Me.Width - 170
   End With
End Sub

