Attribute VB_Name = "Publics"
' Publics.bas


Option Explicit
Option Base 1

Public Const SIF_RANGE = &H1
Public Const SIF_PAGE = &H2
Public Const SIF_POS = &H4
Public Const SIF_DISABLENOSCROLL = &H8
Public Const SIF_TRACKPOS = &H10
Public Const SB_VERT = 1      ' = SIF_RANGE
Public Const SB_BOTTOM = 7
Public Const SIF_ALL = SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS

Public Type SCROLLINFO
        cbSize As Long
        fMask As Long
        nMin As Long
        nMax As Long
        nPage As Long
        NPos As Long
        nTrackPos As Long
End Type

Public Declare Function GetScrollRange Lib "user32" _
   (ByVal hwnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Public Declare Function GetScrollPos Lib "user32" _
   (ByVal hwnd As Long, ByVal nBar As Long) As Long
Public Declare Function GetScrollInfo Lib "user32" _
   (ByVal hwnd As Long, ByVal N As Long, lpScrollInfo As SCROLLINFO) As Long


Const TTEESSTT = 0


' HighLlighting code by buggy PSC CodeId = 43509
' Hide/Show srcollbars code by Andrew Baker  @ www.vbuser.com

Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1
Public Const EM_SCROLL = &HB5
Public Const EM_LINESCROLL = &HB6
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_GETTHUMB = &HBE

Public Const EM_LINELENGTH = &HC1
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_CHARFROMPOS = &HD7
Public Const WM_USER = &H400

Public Const EM_GETCHARFORMAT = (WM_USER + 58)


Public Const LF_FACESIZE = 32
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Public Const CFM_BACKCOLOR = &H4000000
'Public Const CFM_BACKCOLOR = &H4FFFFFF
Public Const SCF_SELECTION = &H1

Public Const LB_SETHORIZONTALEXTENT = &H194

' Wordwrap
'To set wordwrap:
'SendMessageLong RTB.hwnd, EM_SETTARGETDEVICE, 0, 0
'To unset wordWrap:
'SendMessageLong RTB.hwnd, EM_SETTARGETDEVICE, 0, 1
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)

Public Type CHARFORMAT2
    cbSize As Integer    '2
    wPad1 As Integer    '4
    dwMask As Long    '8
    dwEffects As Long    '12
    yHeight As Long    '16
    yOffset As Long    '20
    crTextColor As Long    '24
    bCharSet As Byte    '25
    bPitchAndFamily As Byte    '26
    'szFaceName(0 To LF_FACESIZE - 1) As Byte    ' 58
    szFaceName(0 To LF_FACESIZE - 1) As Byte   ' 58
    wPad2 As Integer    ' 60

' Additional stuff supported by RICHEDIT20
    wWeight As Integer    ' /* Font weight (LOGFONT value)      */
    sSpacing As Integer    ' /* Amount to space between letters  */
    crBackColor As Long    ' /* Background color                 */
    lLCID As Long    ' /* Locale ID                        */
    dwReserved As Long    ' /* Reserved. Must be 0              */
    sStyle As Integer    ' /* Style handle                     */
    wKerning As Integer    ' /* Twip size above which to kern char pair*/
    bUnderlineType As Byte    ' /* Underline type                   */
    bAnimation As Byte    ' /* Animated text like marching ants */
    bRevAuthor As Byte    ' /* Revision author index            */
    bReserved1 As Byte
End Type

Public Type POINTAPI
   kx As Long
   ky As Long
End Type

Public Type RECT
   Left   As Integer
   Top    As Integer
   Right  As Integer
   Bottom As Integer
End Type


Private Declare Function ShowScrollBar Lib "user32" _
   (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
   (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


'----------------------------------------------------
Public FileSpec$, PathSpec$

Public ModuleType$(), NumModTypes() As Long
Public NModTypes As Long
Public ProcName$(), NumProcs() As Long      ' 1- 26
Public ProcSub1 As Long, ProcSub2 As Long, ProcSub3 As Long
Public PubS1 As Long
Public PubE1 As Long
Public PubS2 As Long
Public PubE2 As Long
   ' Private ranges
Public PriS1 As Long
Public PriE1 As Long
Public PriS2 As Long
Public PriE2 As Long

Public VBPDir$
Public ModFileSpec$, ProcTitle$
Public ModString$
Public StartOfProcsPos As Long
Public StartOfCodePos As Long
Public ItemNum As Long
Public ModItemNum As Long

Public SearchText$
Public FoundPos As Long
Public FirstVisLine As Long
Public CurrentVisLine As Long
Public FoundLine As Long
Public PrevFirstVisLine As Long
Public LinesToScroll As Long
Public rtfOptions As Long
Public RTModState As Long  ' 0 nothing, 1 Module or List, 2 Squashed, 3 Stripped
Public aTimer As Boolean
Public LineCount As Long

' Collections
Public NumMods As Long
Public ModName$()
Public ModCtrlName$()
Public MaxNumCtrls As Long
Public ModStartPos() As Long
Public ModProcPos() As Long
Public ModCollection$   ' Whole project stripped
Public RT$  ' Build string for RTMod
Public NumPubPriVars As Long
Public ModNameStore$()
Public PubPrivStore$()

' Listing Dims ReDims & Consts
Public DimStore$()
Public DimReConstType() As Long


Public TotNumArgs As Long
Public NumArgsInProc As Long
Public ArgStore$()
Public AStore$()
Public aBar As Boolean ' To Bar Index, Button, Shift etc
                       ' for unused Proc Arguments
                       
Public aColoringDone As Boolean

Public RTModColor As Long
Public GenFontSize As Long
Public aGenBold As Boolean
Public SomeName$

' frmStats LOCSIZES
Public frmStatsLeft As Long
Public frmStatsTop As Long
Public frmStatsWidth As Long
Public frmStatsHeight As Long

' frmFind LOCSIZES
Public frmFindLeft As Long
Public frmFindTop As Long
Public frmFindWidth As Long
Public frmFindHeight As Long

'Screen.TwipsPerPixelX/Y
Public STX As Long, STY As Long

Public NumEventTypes As Long
Public EventType$()

Public aHelp As Boolean

' Test layouts
Private Enum DD
 DDD ' = 0
End Enum

Type TTT
  tttt As Long
End Type

Const yy = 7.7
Private aa As Long

Public Unuser As Long

    Function GetFileName(FSpec$) _
As _
String

Dim p As Long
   GetFileName = FSpec$
   p = InStrRev(FSpec$, "\")
   If p <> 0 Then
      GetFileName = Right$(FSpec$, Len(FSpec$) - p)
   End If
End Function

Public Sub AddScroll(List As ListBox)
' Public GenFontSize, aGenBold
    Dim i As Long, GreatestLen As Long, GreatestWidth As Long
    'Find Longest Text in Listbox
    GreatestLen = 0
    For i = 0 To List.ListCount - 1
        If Len(List.List(i)) > Len(List.List(GreatestLen)) Then
            GreatestLen = i
        End If
    Next i
    'Get Twips
    GreatestWidth = List.Parent.TextWidth(List.List(GreatestLen) + String$(12, "W"))
    GreatestWidth = GreatestWidth * GenFontSize / 8
    If aGenBold Then
      GreatestWidth = GreatestWidth * 1.2
    End If
    'Space(12) is used to prevent the last Character from being cut off
    'Convert to Pixels if Parent form in Twips
    'GreatestWidth = GreatestWidth \ Screen.TwipsPerPixelX
    'Use api to add scrollbar
    SendMessageLong List.hwnd, LB_SETHORIZONTALEXTENT, GreatestWidth, 0
End Sub

Public Sub FixExtension(FSpec$, Ext$)
' In: FixExtension FileSpec$ & Ext$ (".xxx")
Dim p As Long
   If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   
   p = InStr(1, FSpec$, ".")
   
   If p = 0 Then
      FSpec$ = FSpec$ & Ext$
   Else
      FSpec$ = Mid$(FSpec$, 1, p - 1) & Ext$
   End If
End Sub

Public Function FileExists(ByVal InSpec$) As Boolean
    On Error Resume Next
    FileExists = (Dir$(InSpec$) <> "")
End Function

'Purpose     :  Hides/Shows the Horiz Scroll bar on a list box
'Inputs      :  lbListBox               Listbox to alter
'               bShow                   If True shows Horizontal Scroll bar,
'                                       else hides bar.
'Outputs     :  N/A
'Author      :  Andrew Baker  @ www.vbuser.com
'Date        :  07/10/2000 23:55
'Notes       :
'Revisions   :


Public Sub ListBoxHorScroll(lbListBox As ListBox, bShow As Boolean)
    Const SB_HORZ = 0, SB_CTL = 2, SB_BOTH = 3
    
    If bShow Then
        Call ShowScrollBar(lbListBox.hwnd, SB_HORZ, 1&)
    Else
        Call ShowScrollBar(lbListBox.hwnd, SB_HORZ, 0&)
    End If
End Sub

Public Sub ListBoxVertScroll(lbListBox As ListBox, bShow As Boolean)
    Const SB_VERT = 1, SB_CTL = 2, SB_BOTH = 3
    
    If bShow Then
        Call ShowScrollBar(lbListBox.hwnd, SB_VERT, 1&)
    Else
        Call ShowScrollBar(lbListBox.hwnd, SB_VERT, 0&)
    End If
End Sub

' Also works for RichTextBox

Public Sub RTBHorScroll(RTB As RichTextBox, bShow As Boolean)
    Const SB_HORZ = 0, SB_CTL = 2, SB_BOTH = 3
    If bShow Then
        Call ShowScrollBar(RTB.hwnd, SB_HORZ, 1&)
    Else
        Call ShowScrollBar(RTB.hwnd, SB_HORZ, 0&)
    End If
End Sub

Public Sub FindText(RTB As RichTextBox)
' Public FoundPos
' Public SearchText$
Dim k As Long
   'Label4 = ""
   k = 0
   CurrentVisLine = SendMessageLong(RTB.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
   FoundPos = RTB.Find(SearchText$, k, , rtfOptions)
   If FoundPos <> -1 Then ' SearchText$ string found
   '   On Error Resume Next

      RTB.SetFocus
      FirstVisLine = SendMessageLong(RTB.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
      FoundLine = RTB.GetLineFromChar(FoundPos)
      If FirstVisLine >= CurrentVisLine Then ' Scroll RTMod up
         LinesToScroll = (FoundLine - FirstVisLine)
         SendMessageLong RTB.hwnd, EM_LINESCROLL, 0&, ByVal (LinesToScroll)
      Else
      End If
      RTB.SelStart = FoundPos   ' set selection start and
      RTB.SelLength = Len(SearchText$) + 2 ' set selection length
      RTModColor = vbRed
      RTB.SelColor = vbRed
      RTB.SelLength = 0
   Else 'Not found
   End If
End Sub

Public Sub FindNextText(RTB As RichTextBox)
Dim k As Long
   If Not aColoringDone Then Exit Sub
   'Label4 = ""
   k = FoundPos + 1
   If k < 0 Then Exit Sub
   If k > Len(RTB.Text) Then Exit Sub
   CurrentVisLine = SendMessageLong(RTB.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
   FoundPos = RTB.Find(SearchText$, k, , rtfOptions)
   If FoundPos <> -1 Then ' SearchText$ string found
      'On Error Resume Next
      RTB.SetFocus
      FirstVisLine = SendMessageLong(RTB.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
      FoundLine = RTB.GetLineFromChar(FoundPos)

      If FirstVisLine >= CurrentVisLine Then ' Scroll RTMod up
         LinesToScroll = (FoundLine - FirstVisLine)
         SendMessageLong RTB.hwnd, EM_LINESCROLL, 0&, ByVal (LinesToScroll)
      Else
      End If
      RTB.SelStart = FoundPos   ' set selection start and
      RTB.SelLength = Len(SearchText$) ' set selection length
      RTB.SelColor = RTModColor
      RTB.SelLength = 0
   Else 'Not found
   End If
End Sub

