VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMax 
   Caption         =   " MaxMod"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9345
   LinkTopic       =   "Form2"
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   623
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdFindNext 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   300
      Left            =   5715
      Picture         =   "frmMax.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   " Find Next "
      Top             =   45
      Width           =   360
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0E0FF&
      Height          =   300
      Left            =   5295
      Picture         =   "frmMax.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " Find "
      Top             =   45
      Width           =   390
   End
   Begin VB.CheckBox chkRTB 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   3675
      Picture         =   "frmMax.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   45
      Width           =   345
   End
   Begin VB.CheckBox chkRTB 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   4050
      Picture         =   "frmMax.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   45
      Width           =   345
   End
   Begin VB.ComboBox cboFS 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   315
      ItemData        =   "frmMax.frx":0528
      Left            =   3045
      List            =   "frmMax.frx":052A
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   " Font size "
      Top             =   45
      Width           =   615
   End
   Begin VB.CommandButton cmdColorIt 
      BackColor       =   &H00C0E0FF&
      Height          =   300
      Left            =   4455
      Picture         =   "frmMax.frx":052C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " Color RTB "
      Top             =   45
      Width           =   360
   End
   Begin VB.CommandButton cmdStopColoring 
      BackColor       =   &H00C0E0FF&
      Height          =   300
      Left            =   4830
      Picture         =   "frmMax.frx":0DF6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " Stop coloring "
      Top             =   45
      Width           =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8775
      Top             =   15
   End
   Begin VB.CommandButton cmdHiLit 
      BackColor       =   &H00C0E0FF&
      Height          =   300
      Left            =   6090
      Picture         =   "frmMax.frx":16C0
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Highliight selected text "
      Top             =   45
      Width           =   360
   End
   Begin VB.PictureBox picPB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   6645
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   119
      TabIndex        =   1
      Top             =   75
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox RTBMax 
      Height          =   9030
      Left            =   120
      TabIndex        =   0
      Top             =   405
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   15928
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMax.frx":180A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LabName 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   135
      TabIndex        =   11
      Top             =   120
      Width           =   45
   End
   Begin VB.Label LabLine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Line"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   6930
      TabIndex        =   10
      Top             =   180
      Width           =   1245
   End
   Begin VB.Shape Shape2 
      Height          =   390
      Left            =   3015
      Top             =   15
      Width           =   3510
   End
End
Attribute VB_Name = "frmMax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmMAX.frm

Option Explicit

Private Sub Form_Load()
   aGenBold = False
   cboFS.AddItem "8"
   cboFS.AddItem "10"
   cboFS.AddItem "12"
   cboFS.ListIndex = 0
   cboFS_Click
   'RTBMax wordwrap OFF
   SendMessageLong RTBMax.hwnd, EM_SETTARGETDEVICE, 0, 1

   RTBMax.Text = Form1.RTMod.Text
   RTBMax.Refresh
   LabName = SomeName$
   Caption = SomeName$ & " (Maximized)"
End Sub

Private Sub cboFS_Click()
' Public GenFontSize
Dim i As Long
   If Not aColoringDone Then
      Exit Sub
   End If
   GenFontSize = Val(cboFS.Text)
   RTBMax.Font.Size = GenFontSize
End Sub

Private Sub chkRTB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
' Public aGenBold
   If Not aColoringDone Then
      chkRTB(Index).Value = 1 - chkRTB(Index).Value
      Exit Sub
   End If
   If Index = 0 Then
      aGenBold = -chkRTB(0)
      RTBMax.Font.Bold = aGenBold
   Else
      RTBMax.Font.Underline = -chkRTB(1)
   End If
   '   On Error Resume Next
   RTBMax.SetFocus
End Sub

Private Sub cmdColorIt_Click()
   ColorRTB frmMax, RTBMax
   aColoringDone = True
End Sub

Private Sub cmdStopColoring_Click()
   aColoringDone = True
End Sub

Private Sub cmdHiLit_Click()
'buggy PSC CodeId=43509
Dim RTFformat As CHARFORMAT2
   If Not aColoringDone Then Exit Sub
   If FoundPos > -1 Then
      With RTFformat
          .cbSize = Len(RTFformat)
          .dwMask = CFM_BACKCOLOR
          .crBackColor = vbCyan
      End With
      SendMessage RTBMax.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, RTFformat
      RTBMax.SelStart = FoundPos   ' set selection start
   End If
End Sub

'#### FINDERS #############################################

Private Sub cmdFind_Click()
   ' Get entered search string
   If Not aColoringDone Then Exit Sub
   frmFind.Show vbModal
   If LenB(SearchText$) <> 0 Then
      FoundPos = -1
      RTModColor = vbRed
      cmdFindNext_Click
   End If
End Sub

Private Sub cmdFindNext_Click()
   If Not aColoringDone Then Exit Sub
    FindNextText RTBMax
End Sub


Private Sub RTBMax_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim LN As Long    ' Line number
Dim OFFSET As Long
Dim p As Long
Dim PT As POINTAPI
   If Not aColoringDone Then Exit Sub
   If Button = vbLeftButton Then
      PT.kx = x / STX
      PT.ky = y / STY
      OFFSET = SendMessage(RTBMax.hwnd, EM_CHARFROMPOS, 0&, PT)
      LN = SendMessageLong(RTBMax.hwnd, EM_LINEFROMCHAR, -1&, 0&) + 1
      LabLine = " Line" & Str$(LN)
      LabLine.Refresh
      FoundPos = OFFSET
   End If
End Sub

Private Sub Form_Resize()
Dim w1 As Long
Dim h1 As Long
   If WindowState <> vbMinimized Then
      w1 = frmMax.Width / STX - 30
      h1 = frmMax.Height / STY - RTBMax.Top * 3
      If w1 > 60 Then
      If h1 > 60 Then
         RTBMax.Width = w1
         RTBMax.Height = h1
         RTBMax.Refresh
      End If
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   aColoringDone = True
   RTBMax.Text = ""
End Sub
