VERSION 5.00
Begin VB.Form Frm_SrcEdit 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Token Viewer"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4680
      Width           =   7815
   End
   Begin VB.VScrollBar VScroll 
      Height          =   4935
      Left            =   7800
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Fr_Text 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame1"
      Height          =   2835
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.Label Lbl_item 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   32767
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   720
      End
   End
End
Attribute VB_Name = "Frm_SrcEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Txt_LastItem As Label
Dim Txt_ItemCount&

Dim ItemNext_Left&
Dim ItemNext_Top&

Dim FormBaseHeight&
Dim FormBaseWidth&

Dim FrameBaseHeight&

Dim LineHeight&

Private Sub Form_Load()

Me.Show
'   Height = 0
   FormBaseHeight = 512 + HScroll.Height - 4 'Height
   
'   Width = 0
   FormBaseWidth = 124 + VScroll.Width - 4 'Width

   Fr_Text.Height = 0
   FrameBaseHeight = Fr_Text.Height

   LineHeight = Lbl_item(&H7FFF).Height
   
   
   
'    With Controls.Add("VB.TextBox", "Text1")
'         .Text = "test"
'         .FontBold = True
'         '.Alignment = vbCenter
''         .BackStyle = BorderStyleConstants.vbTransparent
' '        .AutoSize = True
''         .Top = i.Top + i.Height / 3
''         .Left = i.Left + i.Width
'         .Visible = True
''         .Index = 1
'      End With
End Sub

Sub LineBreak()
   With Txt_LastItem
      
      Fr_Text.Width = Max(Fr_Text.Width, ItemNext_Left)
      
      ItemNext_Left = 0
      ItemNext_Top = ItemNext_Top + LineHeight
      
      Fr_Text.Height = ItemNext_Top + LineHeight + FrameBaseHeight _
      
      
     'scroll
      Fr_Text_VScroll
      
   End With
   


End Sub

Sub Fr_Text_VScroll(Optional Percent As Double = 1)
   Dim TopPos&
   TopPos = (Height - Fr_Text.Height - FormBaseHeight)
   If TopPos < 0 Then
      
      Fr_Text.Top = TopPos * Percent
      
  'Enable/Disable bars
      VScroll.Visible = True
   Else
      Fr_Text.Top = 0
      VScroll.Visible = False
   End If
End Sub

Sub Fr_Text_HScroll(Optional Percent As Double = 1)
   Dim LeftPos&
   LeftPos = (Width - Fr_Text.Width - FormBaseWidth)
   If LeftPos < 0 Then
   
      Fr_Text.Left = LeftPos * Percent
      
  'Enable/Disable bars
      HScroll.Visible = True
   Else
      Fr_Text.Left = 0
      HScroll.Visible = False
   End If
End Sub


Function AddItem(Text$, _
         Optional Color&, _
         Optional TypeName$, _
         Optional TokenInfo$, _
         Optional bLineBreak As Boolean = False _
         ) As Label
   
   
   If Txt_ItemCount >= &H7FFE Then Exit Function
   
   Load Lbl_item(Txt_ItemCount)
   Set Txt_LastItem = Lbl_item(Txt_ItemCount)
   Set AddItem = Txt_LastItem
   
   With Txt_LastItem
      
      .Left = ItemNext_Left
      .Top = ItemNext_Top
      
    ' use random based color set
      Rnd -1
      Randomize Color
      .ForeColor = Rnd * &HFFFFFF
      
      
      .Caption = Text
      
      .ToolTipText = TypeName & "[" & TokenInfo & "]"

      
      .Visible = True
      
     'Take care of possible linebreaks(vbcrlf)
      ItemNext_Top = ItemNext_Top + _
                     .Height - LineHeight
      
      If bLineBreak Then
         LineBreak
      Else
       ' Calc next char pos
         ItemNext_Left = .Left + .Width
      End If
      
      Fr_Text.Width = Max(Fr_Text.Width, ItemNext_Left)

      
   End With
   Inc Txt_ItemCount
   
'   Debug.Print Txt_ItemCount
'   Dim i
'   For i = 0 To 100000
   myDoEvents
 '  Next
   
End Function

Private Sub Form_Resize()
On Error Resume Next
   With VScroll
    ' Sync scrollbar height with the form
      .Height = Height - FormBaseHeight
    
    ' Attach scrollbar to the left boarder of the form
      .Left = Width - FormBaseWidth
   End With
   
   With HScroll
    ' Sync scrollbar Width with the form
      .Width = Width - FormBaseWidth
    
    ' Attach scrollbar to the top boarder of the form
      .Top = Height - FormBaseHeight
   End With
   
   Fr_Text_VScroll
   Fr_Text_HScroll
   
 myDoEvents
End Sub

Private Sub HScroll_Change()
   With HScroll
      Dim Percent As Double
      Percent = .value / (.Max - .Min)
      Fr_Text_HScroll Percent
      
   End With
   
End Sub

Private Sub VScroll_Scroll()
   With VScroll
      Dim Percent As Double
      Percent = .value / (.Max - .Min)
      Fr_Text_VScroll Percent
      
   End With
End Sub
