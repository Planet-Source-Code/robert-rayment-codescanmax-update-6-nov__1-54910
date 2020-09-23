VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find text"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3435
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   76
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   229
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRTF 
      Caption         =   "Whole Word"
      Height          =   210
      Index           =   1
      Left            =   1935
      TabIndex        =   4
      Top             =   165
      Width           =   1350
   End
   Begin VB.CheckBox chkRTF 
      Caption         =   "Match Case"
      Height          =   210
      Index           =   0
      Left            =   225
      TabIndex        =   3
      Top             =   165
      Width           =   1320
   End
   Begin VB.ComboBox cboSearch 
      Height          =   315
      Left            =   780
      TabIndex        =   0
      Top             =   495
      Width           =   1860
   End
   Begin VB.CommandButton cmdFindIt 
      Caption         =   "Cancel"
      Height          =   285
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   495
      Width           =   690
   End
   Begin VB.CommandButton cmdFindIt 
      Caption         =   "OK"
      Height          =   285
      Index           =   0
      Left            =   45
      TabIndex        =   2
      Top             =   495
      Width           =   690
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmFind.frm

Option Explicit

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Const hWndInsertAfter = -1 ' On top
'Private Const wFlags = &H1 Or &H2 Or &H40  ' No move & Show window
Private Const wFlags = &H1 Or &H40   ' No move & Show window


Private Sub cmdFindIt_Click(Index As Integer)
Dim N As Long
   
' Set rtfOptions
' Public rtfOptions
' Index              0                1
' rtfOptions = [rtfMatchCase] or [rtfWholeWord]    'Or rtfNotHighlight)
   If chkRTF(0).Value = Checked And chkRTF(1).Value = Checked Then
      rtfOptions = rtfMatchCase Or rtfWholeWord
   ElseIf chkRTF(0).Value = Checked Then
      rtfOptions = rtfMatchCase
   ElseIf chkRTF(1).Value = Checked Then
      rtfOptions = rtfWholeWord
   Else
      rtfOptions = 0
   End If
   
   If Index = 1 Then    ' Cancel
      SearchText$ = ""
   Else
      SearchText$ = cboSearch.Text
   End If
   
   ' AddItem SearchText$ if not already there
   If LenB(SearchText$) <> 0 Then
      If cboSearch.ListCount > 0 Then
         For N = 0 To cboSearch.ListCount - 1
            If SearchText$ = cboSearch.List(N) Then Exit For
         Next N
         If N = cboSearch.ListCount Then
            cboSearch.AddItem SearchText$, (0)
         End If
      Else
         cboSearch.AddItem SearchText$, (0)
      End If
   End If
   'Unload frmFind
   frmFindLeft = frmFind.Left
   frmFindTop = frmFind.Top
   frmFind.Hide   ' Keeps settings
End Sub

Private Sub Form_Activate()
' Public
   'frmFind.Width = 2655
   'frmFind.Height = 8310
   'frmFindLeft = 40
   'frmFindTop = 66
   frmFind.Width = frmFindWidth
   frmFind.Height = frmFindHeight
   
   
   If Form1.WindowState <> vbMaximized Then
      frmFind.Left = Form1.Left - frmFindWidth
      frmFind.Top = Form1.Top
   Else  ' Maximised
      ' Size & Make frmFind stay on top
      SetWindowPos frmFind.hwnd, hWndInsertAfter, frmFindLeft / STX, frmFindTop / STY, _
      frmFindWidth / STX, frmFindHeight / STY, wFlags
   End If
   cboSearch.SetFocus
End Sub
