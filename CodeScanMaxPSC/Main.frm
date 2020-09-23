VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   9390
   DrawWidth       =   2
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   626
   Begin VB.CommandButton cmdMaxRTB 
      BackColor       =   &H00FF8080&
      Height          =   345
      Left            =   8865
      Picture         =   "Main.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   " Max RTB "
      Top             =   2730
      Width           =   390
   End
   Begin VB.PictureBox picPB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   7335
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   119
      TabIndex        =   25
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdHiLit 
      BackColor       =   &H00C0E0FF&
      Height          =   300
      Left            =   6675
      Picture         =   "Main.frx":0F14
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   " Highliight selected text "
      Top             =   2400
      Width           =   360
   End
   Begin VB.CommandButton cmdHSSBS 
      BackColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   1
      Left            =   6270
      Picture         =   "Main.frx":105E
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   " Show Horz SBars "
      Top             =   2400
      Width           =   360
   End
   Begin VB.CommandButton cmdHSSBS 
      BackColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   0
      Left            =   5865
      Picture         =   "Main.frx":15E8
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   " Hide Horz SBars "
      Top             =   2400
      Width           =   360
   End
   Begin VB.CheckBox chkLineNumbers 
      BackColor       =   &H00C0E0FF&
      Height          =   300
      Left            =   5445
      Picture         =   "Main.frx":1B72
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   " Toggle Line Numbers "
      Top             =   2400
      Width           =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7710
      Top             =   2580
   End
   Begin RichTextLib.RichTextBox RTBLN 
      Height          =   3270
      Left            =   15
      TabIndex        =   20
      Top             =   3090
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   5768
      _Version        =   393217
      BackColor       =   -2147483626
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Main.frx":1C44
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
   Begin VB.CommandButton cmdStopColoring 
      BackColor       =   &H00C0E0FF&
      Height          =   300
      Left            =   4155
      Picture         =   "Main.frx":1CC6
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   " Stop coloring "
      Top             =   2400
      Width           =   360
   End
   Begin VB.CommandButton cmdColorIt 
      BackColor       =   &H00C0E0FF&
      Height          =   300
      Left            =   3780
      Picture         =   "Main.frx":2590
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   " Color RTB "
      Top             =   2400
      Width           =   360
   End
   Begin VB.ComboBox cboFS 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   315
      ItemData        =   "Main.frx":2E5A
      Left            =   2370
      List            =   "Main.frx":2E5C
      Style           =   2  'Dropdown List
      TabIndex        =   16
      ToolTipText     =   " Font size "
      Top             =   2400
      Width           =   615
   End
   Begin VB.CheckBox chkRTB 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   3375
      Picture         =   "Main.frx":2E5E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2400
      Width           =   345
   End
   Begin VB.CheckBox chkRTB 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   3000
      Picture         =   "Main.frx":2FA8
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
      Width           =   345
   End
   Begin VB.CommandButton cmdCollect 
      BackColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   1
      Left            =   1680
      Picture         =   "Main.frx":30F2
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   " Strip ALL "
      Top             =   2430
      Width           =   480
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0E0FF&
      Height          =   300
      Left            =   4620
      Picture         =   "Main.frx":3244
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   " Find "
      Top             =   2400
      Width           =   390
   End
   Begin VB.CommandButton cmdCollect 
      BackColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   0
      Left            =   1170
      Picture         =   "Main.frx":338E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   " Squash ALL "
      Top             =   2430
      Width           =   480
   End
   Begin VB.CommandButton cmdFindNext 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   300
      Left            =   5040
      Picture         =   "Main.frx":34E0
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   " Find Next "
      Top             =   2400
      Width           =   360
   End
   Begin VB.CommandButton cmdStrip 
      BackColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   1
      Left            =   630
      Picture         =   "Main.frx":362A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " Strip Module "
      Top             =   2430
      Width           =   480
   End
   Begin VB.CommandButton cmdStrip 
      BackColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   0
      Left            =   120
      Picture         =   "Main.frx":377C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " Squash Module "
      Top             =   2430
      Width           =   480
   End
   Begin RichTextLib.RichTextBox RTMod 
      Height          =   3315
      Left            =   1005
      TabIndex        =   2
      Top             =   3105
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   5847
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Main.frx":38CE
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
   Begin VB.ListBox ListProcs 
      Height          =   2040
      IntegralHeight  =   0   'False
      Left            =   2295
      TabIndex        =   1
      Top             =   255
      Width           =   6930
   End
   Begin VB.ListBox ListMods 
      Height          =   2010
      IntegralHeight  =   0   'False
      Left            =   75
      TabIndex        =   0
      Top             =   255
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      Height          =   390
      Left            =   2325
      Top             =   2370
      Width           =   4770
   End
   Begin VB.Shape Shape1 
      Height          =   390
      Left            =   45
      Top             =   2355
      Width           =   2205
   End
   Begin VB.Label LabVBP 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   75
      TabIndex        =   19
      Top             =   30
      Width           =   45
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2355
      TabIndex        =   13
      Top             =   2820
      Width           =   585
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "RTB"
      Height          =   195
      Left            =   105
      TabIndex        =   8
      Top             =   2820
      Width           =   345
   End
   Begin VB.Label LabVarProc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Decs && Procs"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4950
      TabIndex        =   6
      Top             =   15
      Width           =   990
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6465
      TabIndex        =   4
      Top             =   2835
      Width           =   495
   End
   Begin VB.Label LabSize 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Size"
      Height          =   240
      Left            =   480
      TabIndex        =   3
      Top             =   2775
      Width           =   1845
   End
   Begin VB.Menu mnuF 
      Caption         =   "&FILE"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open file (.vbp,*.*)"
      End
      Begin VB.Menu zbrk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveRTB 
         Caption         =   "&Save RTB As  *.txt only"
      End
      Begin VB.Menu mnuPrintRTB 
         Caption         =   "&Print RTB"
      End
      Begin VB.Menu zbrk2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuStats 
      Caption         =   "&Stats"
      Begin VB.Menu mnuModStats 
         Caption         =   "Module Stats"
      End
      Begin VB.Menu mnuAllModStats 
         Caption         =   "All Modules' Stats"
      End
   End
   Begin VB.Menu mnuLister 
      Caption         =   "&List All A"
      Begin VB.Menu mnuList 
         Caption         =   "Public Sub"
         Index           =   1
      End
      Begin VB.Menu mnuList 
         Caption         =   "Public Function"
         Index           =   2
      End
      Begin VB.Menu mnuList 
         Caption         =   "Public Property"
         Index           =   3
      End
      Begin VB.Menu mnuList 
         Caption         =   "Static Sub"
         Index           =   4
      End
      Begin VB.Menu mnuList 
         Caption         =   "Static Function"
         Index           =   5
      End
      Begin VB.Menu mnuList 
         Caption         =   "Static Property"
         Index           =   6
      End
      Begin VB.Menu mnuList 
         Caption         =   "Private Sub"
         Index           =   7
      End
      Begin VB.Menu mnuList 
         Caption         =   "Private Function"
         Index           =   8
      End
      Begin VB.Menu mnuList 
         Caption         =   "Private Property"
         Index           =   9
      End
      Begin VB.Menu mnuList 
         Caption         =   "Friend Sub"
         Index           =   10
      End
      Begin VB.Menu mnuList 
         Caption         =   "Friend Function"
         Index           =   11
      End
      Begin VB.Menu mnuList 
         Caption         =   "Friend Property"
         Index           =   12
      End
      Begin VB.Menu mnuList 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuList 
         Caption         =   "Public Declare Sub"
         Index           =   14
      End
      Begin VB.Menu mnuList 
         Caption         =   "Public Declare Function"
         Index           =   15
      End
      Begin VB.Menu mnuList 
         Caption         =   "Public Enum"
         Index           =   16
      End
      Begin VB.Menu mnuList 
         Caption         =   "Public Type"
         Index           =   17
      End
      Begin VB.Menu mnuList 
         Caption         =   "Public Const"
         Index           =   18
      End
      Begin VB.Menu mnuList 
         Caption         =   "Public Event"
         Index           =   19
      End
      Begin VB.Menu mnuList 
         Caption         =   "Public WithEvents"
         Index           =   20
      End
      Begin VB.Menu mnuList 
         Caption         =   "Private Declare Sub"
         Index           =   21
      End
      Begin VB.Menu mnuList 
         Caption         =   "Private Declare Function"
         Index           =   22
      End
      Begin VB.Menu mnuList 
         Caption         =   "Private Enum"
         Index           =   23
      End
      Begin VB.Menu mnuList 
         Caption         =   "Private Type"
         Index           =   24
      End
      Begin VB.Menu mnuList 
         Caption         =   "Private Const"
         Index           =   25
      End
      Begin VB.Menu mnuList 
         Caption         =   "Private Event"
         Index           =   26
      End
      Begin VB.Menu mnuList 
         Caption         =   "Private WithEvents"
         Index           =   27
      End
      Begin VB.Menu mnuList 
         Caption         =   "-"
         Index           =   28
      End
      Begin VB.Menu mnuList 
         Caption         =   "Public VARS"
         Index           =   29
      End
      Begin VB.Menu mnuList 
         Caption         =   "Private VARS"
         Index           =   30
      End
   End
   Begin VB.Menu mnuListerB 
      Caption         =   "List All B"
      Begin VB.Menu mnuListProcs_Details 
         Caption         =   "List Procs with  Dims, ReDims && Consts"
         Index           =   0
      End
      Begin VB.Menu mnuListProcs_Details 
         Caption         =   "List Procs with Arguments"
         Index           =   1
      End
      Begin VB.Menu mnuListProcs_Details 
         Caption         =   "List Control Proc Names"
         Index           =   2
      End
      Begin VB.Menu mnuListProcs_Details 
         Caption         =   "List Non-Control Proc Names"
         Index           =   3
      End
      Begin VB.Menu mnuListProcs_Details 
         Caption         =   "List Control Names"
         Index           =   4
      End
      Begin VB.Menu mnuListProcs_Details 
         Caption         =   "List Non-Control Proc Callers"
         Index           =   5
      End
   End
   Begin VB.Menu mnuUnusedA 
      Caption         =   "&Unused A"
      Begin VB.Menu mnuUNUA 
         Caption         =   "Public Declare Sub"
         Index           =   14
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "Public Declare Function"
         Index           =   15
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "Public Enum"
         Index           =   16
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "Public Type"
         Index           =   17
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "Public Const"
         Index           =   18
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "Public Event"
         Index           =   19
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "Public WithEvents"
         Index           =   20
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "Private Declare Sub"
         Index           =   21
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "Private Declare Function"
         Index           =   22
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "Private Enum"
         Index           =   23
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "Private Type"
         Index           =   24
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "Private Const"
         Index           =   25
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "Private Event"
         Index           =   26
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "Private WithEvents"
         Index           =   27
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "-"
         Index           =   28
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "Public VARS"
         Index           =   29
      End
      Begin VB.Menu mnuUNUA 
         Caption         =   "Private VARS"
         Index           =   30
      End
   End
   Begin VB.Menu mnuUnusedB 
      Caption         =   "&Unused B"
      Begin VB.Menu mnuUNUB 
         Caption         =   "Unused Proc Dims"
         Index           =   0
      End
      Begin VB.Menu mnuUNUB 
         Caption         =   "Unused Proc Arguments ALL"
         Index           =   1
      End
      Begin VB.Menu mnuUNUB 
         Caption         =   "Unused Proc Arguments BAR Button && Shift"
         Index           =   2
      End
      Begin VB.Menu mnuUNUB 
         Caption         =   "Unused Non-Control Procs"
         Index           =   3
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Main.frm  Form1

' CodeScanMax by Robert Rayment (Nov 2004) 1
' VB6 only:  InstrRev, Replace used

'6 Nov
'  Allow for multiple ' marks in proc header lines

'1 Aug
'  Work around for forms with no code
'  Address some differences between Win98 & WinXp
'     RTB versions when setting Fonts & Coloring
'''''''''''''''''''''''''''''''''''''''''''''''''''


' Some basic ideas from PSC CodeId=39149
'  'Code Statistics and Unused Variable Finder' by E O'Sullivan.


'1.  LIST PROJFILES (FileSpec$, ModuleType$()) FROM *.vbp file
'2.  LIST PROCS (ModFileSpec$, ProcName$()) IN A PROJFILE
'3.  SHOW SELECTED PROC (ProcTitle$ = selected ProcName$(#))
'4.  List Control names, Control Procs, Non-control Procs, Non-control Proc Callers
'5.  List All Used & Unused Declarations, Proc Dims, Proc Arguments & Non-control Procs

'    NB. Unused only based on > 1 occurrence of item, so gives guidance only!
'    EG  If there is a Private Sub & a Public Sub in a different module, both
'        unused but with exactly the same name, then the Private Sub will be
'        marked unused but not the Public Sub.
'        Also some unused vars need to be kept eg Types for APIs.

'6.  Save all or Copy/Print all or selection from RichTextBox RTB  (ie RTMod)
'7.  Concatenates all stripped mods into ModCollection$.
'    Done at loading project files and kept throughout.
'    Names & offsets to Start of Code & Start of Procs stored.
'8.  Individual stripped mods in ModString$ - extracted from ModCollection$.

' Limitations so far:-
' Assumes variables are defined
' Not done:-
' Unused Controls
' Unused Modules
' Unused Control Procedures
' Only checks vars as declared
'    eg var$ can be used as var without $ then NOT checked
'       or var As String can be used as var$ also NOT checked.


''''' Used :- ''''''''''''''''''''''''''''''''''''''''''''''
'    "String" to bytes:
'    CopyMemory ByteArr(SIndex), ByVal AString$, Len
'    Bytes to "string":
'    AString$ = Space$(Len)
'    CopyMemory ByVal AString$, ByteArr(SIndex), Len
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Option Base 1

'-- For resizing combobox ------------------------------------------------
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
   ByVal x As Long, ByVal y As Long, _
   ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'-------------------------------------------------------------------------


Private FSize As Long

'TEST -----------------
Private Type RR
   RRR As Long
End Type
Enum EE
 EEE = 9
End Enum
Private Const XX = 8.3
Const NNN = 10
'-----------------------

' For Resizing --------------
Private ORGFrmWid As Long
Private ORGFrmHit As Long
Private ORGFontSize As Long

Private Type CScales
  zLef As Single
  zTop As Single
  zWid As Single
  zHit As Single
  zFSize As Single
End Type
Private ScaleArray() As CScales
Private zRTBFS As Single
'----------------------------

Private CommonDialog1 As OSDialog

Private Sub AnUnusedProc()
'
End Sub

Private Sub InitModProcs()    'CCCCCCCCCCCCCCC TEST
' Set up Public variables

'TEST--------------
Dim A$
A$ = " 4444 ' 5555"
Dim uuuu
'------------------


   NModTypes = 6
   ReDim ModuleType$(NModTypes)
   ReDim NumModTypes(NModTypes)
   ModuleType$(1) = "FORM="            ' *.frm
   ModuleType$(2) = "MODULE="          ' *.bas  ' No ctrls
   ModuleType$(3) = "CLASS="           ' *.cls  ' No ctrls
   ModuleType$(4) = "USERCONTROL="     ' *.ctl
   ModuleType$(5) = "DESIGNER="        ' *.dsr  ' No ctrls
   ModuleType$(6) = "PROPERTYPAGE="    ' *.pag
   ' Accessed in mnuOpen_Click()
   
   ' SubScript ranges for filtering.
   ' Need changing if Proc names changed
   '  Also mnuList & mnuUNUA indexes need to match
   ProcSub1 = 30
   ProcSub2 = 33
   ProcSub3 = 40
   ' Public ranges
   PubS1 = 1
   PubE1 = 6
   PubS2 = 14
   PubE2 = 20
   ' Private ranges
   PriS1 = 7
   PriE1 = 12
   PriS2 = 21
   PriE2 = 27
   ReDim ProcName$(ProcSub3)
   ReDim NumProcs(ProcSub3) As Long
   ' Procs
   ProcName$(1) = "Public Sub "           ' PubS1
   ProcName$(2) = "Public Function "
   ProcName$(3) = "Public Property "
   ProcName$(4) = "Static Sub "
   ProcName$(5) = "Static Function "
   ProcName$(6) = "Static Property "      ' PubE1
   
   ProcName$(7) = "Private Sub "          ' PriS1
   ProcName$(8) = "Private Function "
   ProcName$(9) = "Private Property "
   ProcName$(10) = "Friend Sub "
   ProcName$(11) = "Friend Function "
   ProcName$(12) = "Friend Property "     ' PriE1
   
   ProcName$(13) = ""   ' Break
   
   ' Declarations
   ProcName$(14) = "Public Declare Sub "  ' PubS2
   ProcName$(15) = "Public Declare Function "
   ProcName$(16) = "Public Enum "
   ProcName$(17) = "Public Type "
   ProcName$(18) = "Public Const "
   ProcName$(19) = "Public Event "
   ProcName$(20) = "Public WithEvents "   ' PubE2
   
   ProcName$(21) = "Private Declare Sub " ' PriS2
   ProcName$(22) = "Private Declare Function "
   ProcName$(23) = "Private Enum "
   ProcName$(24) = "Private Type "
   ProcName$(25) = "Private Const "
   ProcName$(26) = "Private Event "
   ProcName$(27) = "Private WithEvents "  ' PriE2
   
   ProcName$(28) = ""   ' Break
   
   ProcName$(29) = "Public "
   ProcName$(30) = "Private "             ' ProcSub1  VARS
   
   ' For Line Input from File also need:-
   ' Unconditioned Procs
   ProcName$(31) = "Sub "
   ProcName$(32) = "Function "
   ProcName$(33) = "Property "            ' ProcSub2
   
   ' Unconditioned Declares      ' Public/Private added when conditioned
   ProcName$(34) = "Declare Sub "
   ProcName$(35) = "Declare Function "
   ProcName$(36) = "Enum "
   ProcName$(37) = "Type "
   ProcName$(38) = "Const "
   ProcName$(39) = "Dim "                 ' Private when conditioned
   ProcName$(40) = "Global "              ' ProcSub3  Public when conditioned
   
   ' Some Events:-  {Can be added to)
   ' For checking non-control proc names using
   ' control name_nonevent
   NumEventTypes = 21
   ReDim EventType$(NumEventTypes)
   EventType$(1) = "Click("
   EventType$(2) = "DblClick("
   EventType$(3) = "MouseUp("
   EventType$(4) = "MouseDown("
   EventType$(5) = "MouseMove("
   EventType$(6) = "KeyUp("
   EventType$(7) = "KeyDown("
   EventType$(8) = "KeyPress("
   EventType$(9) = "Load("
   EventType$(10) = "Unload("
   EventType$(11) = "QueryUnload("
   EventType$(12) = "Resize("
   EventType$(13) = "Activate("
   EventType$(14) = "Initialize("
   EventType$(15) = "Terminate("
   EventType$(16) = "GotFocus("
   EventType$(17) = "LostFocus("
   EventType$(18) = "Change("
   EventType$(19) = "Scroll("
   EventType$(20) = "Paint("
   EventType$(21) = "Timer("

'   EventType$(21) = "Show("              ' ?
'   EventType$(22) = "InitProperties("    ' ?
'   EventType$(23) = "ReadProperties("    ' ?
'   EventType$(24) = "WriteProperties("   ' ?
   
   
   ModFileSpec$ = ""
End Sub


Private Sub cmdMaxRTB_Click()
   SomeName$ = Label3.Caption
   frmMax.Show vbModeless
End Sub

Private Sub Form_Load()
' Public FileSpec$
Dim Ext$
Dim p As Long
Dim FilterIndex As Long
   aTimer = False
   Timer1.Enabled = False
   RTBLN.Text = ""
   
   InitResizeArray
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   Caption = "CodeScanMax  " & Date & "  " & Time
   ORGFrmWid = Form1.Width
   ORGFrmHit = Form1.Height
   ORGFontSize = 8
   FileSpec$ = ""

   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"

   Me.Show
   Me.ScaleMode = vbPixels ' See AddScroll in Publics.bas
   Me.Left = Screen.Width / 4.5
   Me.Top = Screen.Height / 16
   
   Cls
   ListMods.Clear
   ListProcs.Clear
   RTMod.Text = ""
   
   InitModProcs
   
   'RTMod wordwrap OFF
   SendMessageLong RTMod.hwnd, EM_SETTARGETDEVICE, 0, 1
   
   Unload frmStats
   NumMods = 0
   
   mnuSaveRTB.Enabled = False
   mnuPrintRTB.Enabled = False
   
   mnuStats.Enabled = False
   mnuLister.Enabled = False
   mnuListerB.Enabled = False
   mnuUnusedA.Enabled = False
   mnuUnusedB.Enabled = False
   
   cmdStrip(0).Enabled = False
   cmdStrip(1).Enabled = False
   cmdFind.Enabled = False
   cmdFindNext.Enabled = False
   cmdCollect(0).Enabled = False
   cmdCollect(1).Enabled = False
   
   frmStatsWidth = 2655
   frmStatsHeight = 8700
   frmStatsLeft = 40
   frmStatsTop = 66
   
   frmFindWidth = 3525
   frmFindHeight = 1530
   frmFindLeft = 40
   frmFindTop = 66

   GenFontSize = 8
   aGenBold = False
   cboFS.AddItem "8"
   cboFS.AddItem "10"
   cboFS.AddItem "12"
   cboFS.ListIndex = 0
   cboFS_Click
   
   Label3 = ""
   Label4 = ""

   If Command$ <> "" Then
      If FileExists(Command$) Then
         FileSpec$ = Command$
         If LenB(FileSpec$) <> 0 Then
            p = InStrRev(FileSpec$, ".")
            If p <> 0 Then
               Ext$ = LCase$(Right$(FileSpec$, 3))
               If Ext$ = "vbp" Then
                  FilterIndex = 1
               Else
                  FilterIndex = 2
               End If
               MaxNumCtrls = 1
               READFILE FileSpec$, FilterIndex
            End If
         End If
      End If
   ElseIf FileExists(PathSpec$ & "CSInfo.txt") Then
      Open PathSpec$ & "CSInfo.txt" For Input As #1
      Line Input #1, FileSpec$
      Close
      If LenB(FileSpec$) <> 0 Then
         p = InStrRev(FileSpec$, ".")
         If p <> 0 Then
            Ext$ = LCase$(Right$(FileSpec$, 3))
            If Ext$ = "vbp" Then
               FilterIndex = 1
            Else
               FilterIndex = 2
            End If
            MaxNumCtrls = 1
            READFILE FileSpec$, FilterIndex
         End If
      End If
   End If

   InitKeyWords
   aColoringDone = True
   PrevFirstVisLine = -1
End Sub

Private Sub mnuHelp_Click()
   If Not aColoringDone Then Exit Sub
   If Not FileExists(PathSpec$ & "CSHelp.txt") Then
      MsgBox " CSHelp.txt - missing ", vbInformation, "Loading Help file"
      aHelp = False
   Else
      frmHelp.Show vbModeless
   End If
End Sub

Private Sub mnuOpen_Click()
Dim p As Long
Dim Title$, Filt$, InDir$
Dim FIndex As Long

'TEST -----------
Dim UNU(4)
'---------------

   If Not aColoringDone Then Exit Sub

   NumMods = 0

Set CommonDialog1 = New OSDialog

      Title$ = "Open Project File or Any"
      Filt$ = "Open vbp (*.vbp)|*.vbp|All files (*.*)|*.*"
      If FileSpec$ = "" Then
         InDir$ = PathSpec$
      Else
         p = InStrRev(FileSpec$, "\")
         InDir$ = Left$(FileSpec$, p)
         'InDir$ = FileSpec$
      End If
      CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
      ' FIndex = 1 *.vbp file
      ' FIndex = 2 All files *.*
Set CommonDialog1 = Nothing

   MaxNumCtrls = 1
   READFILE FileSpec$, FIndex

End Sub

Private Sub READFILE(FSpec$, FilIndex As Long)
' Called from Form_Load (CSInfo.txt)
'          or mnuOpen_Click
Dim A$, B$
Dim p As Long
Dim k As Long
Dim NU As Long
   On Error GoTo FERR
   Refresh
   Screen.MousePointer = vbHourglass
   DoEvents
   
   If LenB(FSpec$) > 0 Then
      ' Get modules listed in vbp file
      'frmStats.Cls
      RTBLN.Text = ""
      RTMod.Text = ""
      Timer1.Enabled = False
      Unload frmStats
      InitModProcs
      Cls
      CurrentY = 4
      'Print "  " & FSpec$
      If Len(FSpec$) > 44 Then
         LabVBP = Left$(FSpec$, 20) & String$(4, ".") & Right$(FSpec$, 20)
      Else
         LabVBP = FSpec$
      End If
      ListMods.Clear
      ListProcs.Clear
      RTMod.Text = ""
      
      p = InStrRev(FSpec$, "\")
      If p = 0 Then GoTo FERR
      VBPDir$ = Left$(FSpec$, p)
      
      If FilIndex = 1 Then
         If UCase$(Right$(FSpec$, 3)) = "VBP" Then
            Open FSpec$ For Input As #1
            Do Until EOF(1)
               Line Input #1, B$
               A$ = Trim$(UCase$(B$))
               For k = 1 To NModTypes
                  NU = 0
                  p = InStr(1, A$, ModuleType$(k))
                  If p = 1 Then    'Line starting with ModType
                     p = InStr(1, A$, ";")
                     If p = 0 Then
                        p = InStr(1, A$, "=")
                        If p = 0 Then
                           Screen.MousePointer = vbDefault
                           MsgBox A$ & vbCrLf & " Can't understand this line ", vbCritical, "Parsing vbp file"
                           End
                        End If
                     End If
                     B$ = Trim$(Mid$(B$, p + 1))   ' Name
                     A$ = UCase$(Right$(B$, 3))    ' Ext
                     Select Case A$
                     Case "FRM", "BAS", "CLS", "CTL", "DSR", "PAG"
                        NU = NU + 1
                        ListMods.AddItem B$
                     Case "TXT"     ' ie Module=File.txt ??
                        Screen.MousePointer = vbDefault
                        MsgBox A$ & vbCrLf & "file will be ignored ", vbInformation, "Parsing vbp file"
                        Screen.MousePointer = vbHourglass
                     Case Else
                        Screen.MousePointer = vbDefault
                        MsgBox A$ & vbCrLf & " Can't understand this line ", vbCritical, "Parsing vbp file"
                        Screen.MousePointer = vbHourglass
                     End Select
                     Exit For
                  End If
               Next k   ' Test next module type
               If NU > 0 Then NumModTypes(k) = NumModTypes(k) + NU
            Loop  ' Get next line in vbp file
            Close
            
            ' Collect mod file names
            NumMods = ListMods.ListCount
            ReDim ModName$(NumMods)
            ReDim ModCtrlName$(NumMods, 1)
            ReDim ModStartPos(NumMods)
            ReDim ModProcPos(NumMods)
            For k = 0 To NumMods - 1
               ModName$(k + 1) = ListMods.List(k)
            Next k
         End If
         
         ' CONCATENATE all stripped modules
         ' into ModCollection$
         CONCATENATE 1  ' 1 Strips headers
         
         cmdCollect(0).Enabled = True
         cmdCollect(1).Enabled = True
         
         mnuSaveRTB.Enabled = False
         mnuPrintRTB.Enabled = False
         mnuStats.Enabled = True
         mnuModStats.Enabled = False
         mnuLister.Enabled = True
         mnuListerB.Enabled = True
         mnuUnusedA.Enabled = True
         mnuUnusedB.Enabled = True
         cmdFind.Enabled = False
         cmdFindNext.Enabled = False

      Else  ' Any file, FilIndex = 2
         NumMods = 0
         frmStats.Cls
         Unload frmStats
         ModFileSpec$ = ""
         Open FileSpec$ For Binary Access Read As #1
         ModString$ = Space$(LOF(1))
         Get #1, , ModString$
         Close
         RTMod.Text = ModString$
         RTMod.Refresh
         ModString$ = ""
         
         mnuSaveRTB.Enabled = True
         mnuPrintRTB.Enabled = True
         
         mnuStats.Enabled = False
         mnuModStats.Enabled = False
         mnuLister.Enabled = False
         mnuListerB.Enabled = False
         mnuUnusedA.Enabled = False
         mnuUnusedB.Enabled = False
         
         cmdStrip(0).Enabled = False
         cmdStrip(1).Enabled = False
         cmdCollect(0).Enabled = False
         cmdCollect(1).Enabled = False
         
         cmdFind.Enabled = True
         cmdFindNext.Enabled = True
         RTBSizeLines
         DoEvents
      End If
   End If
   AddScroll ListMods
   
   Label3 = ""
   Label4 = ""
   RTModState = 0    ' Nothing
   
   Timer1.Enabled = False
   aTimer = False
   RTBLN.Text = ""
   chkLineNumbers.Value = 0
   
   Screen.MousePointer = vbDefault
Exit Sub
'=========
FERR:
   MsgBox "Can't read the file  " & FileSpec$, vbCritical, "CodeScan"
   Close
   NumMods = 0
   ModString$ = ""
   Label3 = ""
   Label4 = ""
   RTModState = 0    ' Nothing
   On Error GoTo 0
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuAllModStats_Click()
' All Mod Stats
   If Not aColoringDone Then Exit Sub
   Unload frmStats
   If NumMods = 0 Then Exit Sub
   If LenB(FileSpec$) = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   LIST_PROCS_MODCOLLECTION
   If LenB(ModCollection$) = 0 Then
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   ModFileSpec$ = "All Modules"
   LabVarProc = " " & ModFileSpec$ & " - Decs && Procs "
   frmStats.Show vbModeless
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuModStats_Click()
' Selected Mod Stats
   If Not aColoringDone Then Exit Sub
   If NumMods = 0 Then Exit Sub
   If LenB(FileSpec$) > 0 Then
      If ModItemNum >= 0 Then
         Unload frmStats
         ListMods_Click    ' Fills ListProcs & LabVarProc
         ModFileSpec$ = ListMods.List(ModItemNum) ' This is the mod string selected
         GET_MODSTATS ModFileSpec$
         frmStats.Show vbModeless
      End If
   End If
   RTBSizeLines
End Sub

Private Sub ListMods_Click()
' Called from mnuModStats as well as Click event
' List Declaration ProcName$() in selected module
   If Not aColoringDone Then Exit Sub
   Unload frmStats
   ModItemNum = ListMods.ListIndex     ' This is the listbox item number selected
   If ModItemNum < 0 Then Exit Sub
   ModFileSpec$ = ListMods.List(ModItemNum) ' This is the string selected
   LIST_MODULE ModFileSpec$
  ' Show whole ModFileSpec$ in RTMod - from binary input
   Open VBPDir$ & ModFileSpec$ For Binary Access Read As #1
      FSize = LOF(1)
      If FSize = 0 Then
         Close
         MsgBox "Can't read " & VBPDir$ & ModFileSpec$, vbCritical, " Reading VB Files"
         Exit Sub
      End If
      ModString$ = Space$(FSize)
      Get #1, , ModString$
   Close
   RTMod.Text = ""
   RTModColor = 0
   RTMod.SelColor = RTModColor
   RTMod.Text = ModString$
   RTMod.Refresh
   RTModState = 1 ' Module or List
   RTBSizeLines
   ModString$ = ""
   Label3 = " " & ModFileSpec$ & " "
   DoEvents
End Sub

Private Sub LIST_MODULE(TheModSpec$)
' Get module file and show in ListProcs
' Called from:
' ListMods_Click
' cmdStrip_Click
Dim A$, B$
Dim p As Long
Dim k As Long
   If Not aColoringDone Then Exit Sub
   On Error GoTo MODERR
   Label3 = ""
   ListProcs.Clear
   LabVarProc = " " & TheModSpec$ & " - Decs && Procs "
   ' Find Start of Code (ie skip header) - open for sequential Line Input
   Open VBPDir$ & TheModSpec$ For Input As #1
      k = 1
      StartOfCodePos = 1
      Do Until EOF(1)
         Line Input #1, A$
         p = InStr(1, A$, "Attribute VB_")
         If p > 0 Then
            Do Until EOF(1)
               StartOfCodePos = Seek(1)
               Line Input #1, A$
               k = InStr(1, A$, "Attribute VB_")
               If k = 0 Then Exit Do
            Loop
            If k = 0 Then Exit Do
         End If
      Loop
      
      ' List Declares & Find StartOfProcsPos
      Seek #1, StartOfCodePos
      Do Until EOF(1)
         StartOfProcsPos = Seek(1)
         Line Input #1, A$
         B$ = Trim$(A$)
         If B$ <> "" And Left$(B$, 1) <> "'" Then
            ' Test for First Proc ie StartOfProcsPos1
            For k = 1 To PriE1  ' Public Sub -- Friend Property
               p = InStr(1, B$, ProcName$(k))
               If p = 1 Then Exit For
            Next k
            If k < PriE1 + 1 Then Exit Do   ' with StartOfProcsPos
            ' Test for Procs ie StartOfProcsPos2
            For k = ProcSub1 + 1 To ProcSub2 ' Sub,Function,Property
               p = InStr(1, B$, ProcName$(k))
               If p = 1 Then Exit For
            Next k
            If k < ProcSub2 + 1 Then Exit Do  ' with StartOfProcsPos
            ' Test for Declarations1 & to ListProcs
            For k = PubS2 To ProcSub1  ' Public Declare Sub ,,, Private WithEvents,, Public,Private
               If ProcName$(k) <> "" Then ' Break
                  p = InStr(1, B$, ProcName$(k))
                  If p = 1 Then
                     ListProcs.AddItem A$
                     If Right$(A$, 1) = "_" Then
                        Do
                           If Right$(A$, 1) = "_" Then
                              Line Input #1, A$
                              A$ = Space$(6) & A$
                              ListProcs.AddItem A$
                           End If
                        Loop Until Right$(A$, 1) <> "_"
                     End If
                     Exit For
                  End If
               End If
            Next k
            If k = ProcSub1 + 1 Then      ' Not found
               ' Test for Declarations2 & to ListProcs
               For k = ProcSub2 + 1 To ProcSub3    ' Declare Sub,,,Global
                  p = InStr(1, B$, ProcName$(k))
                  If p = 1 Then
                     ListProcs.AddItem A$
                     If Right$(A$, 1) = "_" Then
                        Do
                           If Right$(A$, 1) = "_" Then
                              Line Input #1, A$
                              A$ = Space$(6) & A$
                              ListProcs.AddItem A$
                           End If
                        Loop Until Right$(A$, 1) <> "_"
                     End If
                     Exit For
                  End If
               Next k
            End If
         
         End If   ' If B$ <> "" And Left$(B$, 1) <> "'" Then
      Loop
      
      ' List Procs
      A$ = String$(10, "=")
      ListProcs.AddItem A$
      Seek (1), StartOfProcsPos
      Do Until EOF(1)
         Line Input #1, A$
         B$ = Trim$(A$)
         For k = 1 To PriE1  ' Public Sub -- Friend Property
            p = InStr(1, B$, ProcName$(k))
            If p = 1 Then
               ListProcs.AddItem A$
               If Right$(A$, 1) = "_" Then
                  Do
                     If Right$(A$, 1) = "_" Then
                        Line Input #1, A$
                        A$ = Space$(6) & A$
                        ListProcs.AddItem A$
                     End If
                  Loop Until Right$(A$, 1) <> "_"
               End If
            End If
         Next k
         For k = ProcSub1 + 1 To ProcSub2 ' Sub,Function,Property
            p = InStr(1, B$, ProcName$(k))
            If p = 1 Then
               ListProcs.AddItem A$
               If Right$(A$, 1) = "_" Then
                  Do
                     If Right$(A$, 1) = "_" Then
                        Line Input #1, A$
                        A$ = Space$(6) & A$
                        ListProcs.AddItem A$
                     End If
                  Loop Until Right$(A$, 1) <> "_"
               End If
            End If
         Next k
      
      Loop
   Close
   AddScroll ListProcs
   ModString$ = ""
   
   mnuModStats.Enabled = True
   cmdStrip(0).Enabled = True
   cmdStrip(1).Enabled = True
   
   mnuSaveRTB.Enabled = True
   mnuPrintRTB.Enabled = True
   cmdFind.Enabled = True
   cmdFindNext.Enabled = True
   cmdCollect(0).Enabled = True
   cmdCollect(1).Enabled = True
   Exit Sub
'==============
MODERR:
   MsgBox "Can't read this module file", vbCritical, "CodeScan"
   Close
   ModString$ = ""
   On Error GoTo 0
End Sub

Private Sub ListProcs_Click()
' Find Variable or Proc in RTMod from
'  ListProcs' selected line

' Public SearchText$
Dim p As Long
Dim k As Long
   If Not aColoringDone Then Exit Sub
   ItemNum = ListProcs.ListIndex     ' This is the listbox item number selected
   ProcTitle$ = ListProcs.List(ItemNum) ' This is the string selected
   ' Need to cope with original mod
   ' and stripped down mod.
   ' STRIP out EOL comments & extra spaces
   ProcTitle$ = Trim$(ProcTitle$)
   p = InStrRev(ProcTitle$, "'")
   If p = 1 Then Exit Sub
   If p > 1 Then
      ''' To allow for eg () ' ccccc's
      Do
         ProcTitle$ = Left$(ProcTitle$, p - 1)
         ' Trim any leading spaces & spaces before '
         ProcTitle$ = Trim$(ProcTitle$)
         p = InStrRev(ProcTitle$, "'")
      Loop Until p = 0
      '''''''''''''''''
   End If
   ' Ditch any continuation characters in first line
   If Right$(ProcTitle$, 2) = " _" Then
      ProcTitle$ = Left$(ProcTitle$, Len(ProcTitle$) - 2)
   End If
   ' Public SearchText$
   SearchText$ = ProcTitle$
   rtfOptions = rtfWholeWord Or rtfMatchCase   'Or rtfNotHighlight)
   Label4 = ""
   FindText RTMod   ' SearchText$   'ProcTitle$
   
   If FoundPos <> -1 Then ' SearchText$ string found
      '   On Error Resume Next
      Label4 = " Line" & Str$(FoundLine + 1) & "  Offset" & Str$(FoundPos) & " "
      ' Find Module if RTMOD STRIPPED ALL
      If RTModState = 3 Then GetModName FoundPos ' Name to Label3 (3 Stripped)
      mnuSaveRTB.Enabled = True
      mnuPrintRTB.Enabled = True
   Else 'Not found
      Label4 = " Not found "
   End If
End Sub

'#### FINDERS #############################################

Private Sub cmdFind_Click()
' Public FoundPos
' Public SearchText$
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
' Public FoundPos
' Public SearchText$
' To jump past any repeated entries
' ProcTitle$ & FoundPos from ListProcs
   Label4 = ""

    FindNextText RTMod
   
   If FoundPos <> -1 Then ' SearchText$ string found
      Label4 = " Line" & Str$(FoundLine + 1) & "  Offset" & Str$(FoundPos) & " "
'      ' Find Module if RTMOD STRIPPED ALL
      If RTModState = 3 Then GetModName FoundPos  ' Name to Label3 (3 Stripped)
      mnuSaveRTB.Enabled = True
      mnuPrintRTB.Enabled = True
   Else 'Not found
      Label4 = " None or No more Found "
      FoundPos = -1
      RTMod.SelStart = 0
   End If
End Sub

Private Sub RTMod_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim LN As Long    ' Line number
Dim OFFSET As Long
Dim p As Long
Dim PT As POINTAPI
   If Not aColoringDone Then Exit Sub
   If Button = vbLeftButton Then
      PT.kx = x / STX
      PT.ky = y / STY
      OFFSET = SendMessage(RTMod.hwnd, EM_CHARFROMPOS, 0&, PT)
      LN = SendMessageLong(RTMod.hwnd, EM_LINEFROMCHAR, -1&, 0&) + 1
      Label4 = " Line" & Str$(LN) & "  Offset " & Str$(OFFSET)
      Label4.Refresh
      FoundPos = OFFSET
      ' Find Module if RTMOD STRIPPED ALL
      If RTModState = 3 Then GetModName OFFSET     ' Name to Label3 (3 Stripped)
   End If
End Sub

Private Sub RTBSizeLines()
'Dim LineCount As Long
   LineCount = SendMessageLong(RTMod.hwnd, EM_GETLINECOUNT, 0&, 0&)
   LabSize = Str$(Len(RTMod.Text)) & " B " & Str$(LineCount) & " Lines"
End Sub

Private Sub GetModName(OFFSET As Long)
' Find Module if RTMOD STRIPPED ALL
' Called from:
' cmdFindNext & FindNext
Dim NM As Long
   For NM = 1 To NumMods
         If NM < NumMods Then
            If OFFSET + 1 < ModStartPos(NM + 1) Then
               Label3 = " " & ModName$(NM) & " "
               Exit For
            End If
         Else
            Label3 = " " & ModName$(NM) & " "
            Exit For
         End If
   Next NM
   If NM = NumMods + 1 Then Label3 = ""
End Sub

'#####################################################################

Private Sub cmdStrip_Click(Index As Integer)
' Index = 0 Squash
' Index = 1 STRIP
Dim A$
   If Not aColoringDone Then Exit Sub
   If NumMods = 0 Then Exit Sub
   ModFileSpec$ = ""
   ModItemNum = ListMods.ListIndex     ' This is the listbox item number selected
   If ModItemNum < 0 Then Exit Sub
   ModFileSpec$ = ListMods.List(ModItemNum) ' This is the string selected
   RTMod.Text = ""
   RTModColor = 0
   RTMod.SelColor = RTModColor
   If Index = 0 Then ' Squash
      STRIP ModItemNum + 1, ModFileSpec$, Index
      LabVarProc = " " & ModFileSpec$ & " - Decs && Procs "
      Label3 = " Squashed " & ModFileSpec$ & " "
      RTModState = 1 ' Module or List
      EXTRACT_DECS_PROCS ModFileSpec$
      ' Extract Decs & Procs from ModString$ to ListProcs.AddItem, Name=ModFileSpec$
      A$ = "SQUASHED " & ModFileSpec$
      RTMod.Text = A$ & vbNewLine
   Else     ' STRIP
      EXTRACT_MODSTRING ModFileSpec$       ' Extract ModString$ from ModCollection$
      LabVarProc = " " & ModFileSpec$ & " - Decs && Procs "
      Label3 = " Stripped " & ModFileSpec$ & " "
      RTModState = 1 ' Module or List
      EXTRACT_DECS_PROCS ModFileSpec$
      A$ = "STRIPPED " & ModFileSpec$
      RTMod.Text = A$ & vbNewLine
   End If
   
   ' OUT:  Result in ModString$
   mnuSaveRTB.Enabled = True
   mnuPrintRTB.Enabled = True
   RTMod.Text = RTMod.Text & ModString$
   ColorHeader A$
   RTMod.Refresh
   RTBSizeLines
   
   ListProcs.SetFocus
   AddScroll ListProcs
   '   On Error Resume Next

   RTMod.SetFocus

   DoEvents
End Sub

Private Sub EXTRACT_DECS_PROCS(MName$)
' Extract Decs & Procs from ModString$ to ListProcs.AddItem, Name=ModFileSpec$
' IN ModName$, ModString$, StartOfProcsPos
' Called from cmdStrip  Squash or Strip
Dim A$
Dim pSOL As Long
Dim pEOL As Long
Dim k As Long
Dim FirstIn As Long
Dim p1 As Long

   ListProcs.Clear
   If LenB(ModString$) > 0 Then
      pSOL = 1
      ListProcs.AddItem MName$
      ListProcs.AddItem ""
      FirstIn = 0
      Do
         pEOL = InStr(pSOL, ModString$, Chr$(10))
         If pEOL - pSOL - 1 < 1 Then Exit Do
         ' Collect line
         A$ = Mid$(ModString$, pSOL, pEOL - pSOL - 1)
         If pSOL < StartOfProcsPos Then
            p1 = InStr(A$, "Public")
            If p1 <> 1 Then p1 = InStr(A$, "Private")
            If p1 <> 1 Then p1 = InStr(A$, "Enum")
            If p1 <> 1 Then p1 = InStr(A$, "Type")
            If p1 <> 1 Then p1 = InStr(A$, "Const")
            If p1 <> 1 Then p1 = InStr(A$, "Dim")
            If p1 <> 1 Then p1 = InStr(A$, "Global")
            If p1 = 1 Then
               ListProcs.AddItem A$
            End If
         Else
            For k = 1 To PriE1 ' Public Sub -- Friend Property
               If InStr(1, A$, ProcName$(k)) = 1 Then
                  If FirstIn = 0 Then
                     ListProcs.AddItem "==========="
                     FirstIn = FirstIn + 1
                  End If
                  ListProcs.AddItem A$
                  Exit For
               End If
            Next k
            If k = PriE1 + 1 Then ' Not found yet
               If FirstIn = 0 Then
                  ListProcs.AddItem "==========="
                  FirstIn = FirstIn + 1
               End If
               p1 = InStr(A$, "Sub")
               If p1 <> 1 Then p1 = InStr(A$, "Function")
               If p1 <> 1 Then p1 = InStr(A$, "Property")
               If p1 = 1 Then
                  ListProcs.AddItem A$
               End If
            End If
         End If
         pSOL = pEOL + 1
         If pSOL >= Len(ModString$) Then Exit Do
      Loop   ' Loop thru all lines of ModString$ seeking Decs & Procs
   End If
End Sub

Private Sub cmdCollect_Click(Index As Integer)
' 0 Squash
' 1 STRIP
' CONCATENATE all squashed/stripped modules
   If Not aColoringDone Then Exit Sub
   Screen.MousePointer = vbHourglass
   If Index = 0 Then ' Squash
      CONCATENATE 0     ' Calls STRIP 0 Keep headers. To ModCollection$
      Label3 = " All Mods Squashed "
      RTModState = 2    ' Squashed
   Else ' Also Done @ mnuOpen
      CONCATENATE 1     ' Call STRIP 1 STRIP off header. To ModCollection$
      Label3 = " All Mods Stripped "
      RTModState = 3    ' Stripped
   End If
   If LenB(ModCollection$) = 0 Then Exit Sub
   RTMod.Text = ""
   RTModColor = 0
   RTMod.SelColor = RTModColor
   RTMod.Text = ModCollection$
   RTMod.Refresh
   RTBSizeLines
   mnuSaveRTB.Enabled = True
   mnuPrintRTB.Enabled = True
   ' To ListProcs: Declarations & Procedures
   ' & Count in NumProcs()
   LIST_PROCS_MODCOLLECTION
   LabVarProc = " All modules - Decs && Procs "
   If Index = 0 Then   ' Redo Squashed to Stripped
      CONCATENATE 1    ' Stripped ModCollection$ needed for other scans
   End If
   cmdFind.Enabled = True
   cmdFindNext.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub CONCATENATE(NH As Integer)
' Called from:
' mnuOpen_Click
' cmdStrip_Click Index

' STRIP & CONCATENATE all modules  No header

' NH=0 keep header in STRIP
' NH=1 STRIP header in STRIP

' Called from:
' mnuOpen_Click
' cmdCollect_Click

' Fill:-
' ReDim ModName$(NumMods)
' ReDim ModStartPos(NumMods)
' ReDim ModProcPos(NumMods)
' ModCollection$
' & ModName$() filled
' @ mnuOpen

Dim k As Long
   
On Error GoTo CONCERROR
   If NumMods = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   ModCollection$ = ""
   For k = 1 To NumMods
      ModString$ = ""
      STRIP k, ModName$(k), NH   ' 1 also STRIP off header
      ' OUT:
      ' Result in ModString$  = stripped ModFileSpec$ file contents
      ' StartOfCodePos   eg Option Explicit
      ' StartOfProcsPos   eg Private Sub Form_Load(..
      If Len(ModString$) > 0 Then
         ModStartPos(k) = Len(ModCollection$) + StartOfCodePos
         ModProcPos(k) = Len(ModCollection$) + StartOfProcsPos
         ModCollection$ = ModCollection$ + ModString$
      Else
         'ModCollection$ = ""
      End If
   Next k
   Screen.MousePointer = vbDefault
   Dim N As Long
   N = MaxNumCtrls
Exit Sub
'===============
CONCERROR:
   MsgBox "CONCATENATE error - fatal @ " & ModName$(k), vbCritical, "CONCATENATE"
   Form_Unload 0
   End
End Sub

'#### DIMS, REDIMS & CONSTS ################################################

Private Sub mnuListProcs_Details_Click(Index As Integer)
' List All B
Dim A$
Dim NM As Long
Dim NProcs As Long
Dim NCtrls As Long
Dim TotNumDims As Long
Dim TOTALUNUSED As Long

   If Not aColoringDone Then Exit Sub
   Screen.MousePointer = vbHourglass
   DoEvents

   mnuSaveRTB.Enabled = True
   mnuPrintRTB.Enabled = True
   cmdFind.Enabled = True
   cmdFindNext.Enabled = True
   RTMod.Text = ""
   ListProcs.Clear
   LabVarProc = " Project Files "
   For NM = 1 To NumMods
      ListProcs.AddItem ModName$(NM)
   Next NM
   Select Case Index
   Case 0
      A$ = "PROCS WITH Dims, ReDims & Consts"
      RTMod.Text = A$ & vbNewLine
      LIST_DRCs_MODCOLLECTION 0, NProcs, TotNumDims, TOTALUNUSED   ' LISTTYPE = 0 ' List Procs with Dims ReDims & Consts
      Label3 = Str$(NProcs) & " Procs with " & Str$(TotNumDims) & " Dims, ReDims && Consts "
   Case 1
      aBar = False    ' Allow Button, Shift
      A$ = "PROCS WITH ARGUMENTS"
      RTMod.Text = A$ & vbNewLine
      LIST_ARGS_MODCOLLECTION 1, NProcs, TOTALUNUSED      ' LISTTYPE = 1 List Procs & Arguments
      Label3 = Str$(NProcs) & " Procs with Arguments "
   Case 2
      A$ = "CONTROL PROC NAMES"
      RTMod.Text = A$ & vbNewLine
      LIST_PROC_CALLERS 0, NProcs, TOTALUNUSED
      Label3 = " Control Proc Names" & Str$(NProcs) & " Procs "
   Case 3
      A$ = "NON-CONTROL PROC NAMES"
      RTMod.Text = A$ & vbNewLine
      LIST_PROC_CALLERS 1, NProcs, TOTALUNUSED
      Label3 = " Non-Control Proc Names" & Str$(NProcs) & " Procs "
   Case 4
      A$ = "CONTROL NAMES"
      RTMod.Text = A$ & vbNewLine
      LIST_CTRL_NAMES NCtrls
      Label3 = Str$(NCtrls) & " Controls "
   Case 5
      A$ = "NON-CONTROL PROC CALLERS"
      RTMod.Text = A$ & vbNewLine
      LIST_PROC_CALLERS 2, NProcs, TOTALUNUSED
      Label3 = " Non-Control Proc Callers "
   Case 6
      ' Called from mnuUNUB_Click 6
      RTMod.Text = ""
      A$ = "UNUSED (or Class/Dsr/REv IT) NON-CONTROL PROCS"
      RTMod.Text = A$ & vbNewLine
      LIST_PROC_CALLERS 3, NProcs, TOTALUNUSED          ' List Unused Non-Control Procs
      Label3 = Str$(TOTALUNUSED) & "  Unused (or Class/Dsr/REv IT) Non-Control Procs "
   End Select
   RTMod.Text = RTMod.Text & RT$ '''''''''''''''''''
   ColorHeader A$
   RTMod.Refresh
   
   
'   ' Optional
'   If Index = 5 or Index = 6 Then
'      ' Colorize
'      SearchText$ = " ##"
'      RTModColor = QBColor(9)
'      LockWindowUpdate RTMod.hwnd
'      Do
'         cmdFindNext_Click
'      Loop Until FoundPos < 1
'      RTModColor = vbRed
'      RTMod.SelColor = 0
'      LockWindowUpdate 0
'   End If
   
   RTModState = 1 ' Module or List
   RTBSizeLines
   RT$ = ""
   
   AddScroll ListProcs

   Screen.MousePointer = vbDefault
   DoEvents

End Sub

Private Sub ColorHeader(Search$)
   If InStr(1, Search$, Chr$(10)) Then
      Search$ = Left$(Search$, Len(Search$) - 2)
   End If
   SearchText$ = Search$
   FoundPos = -1
   RTModColor = QBColor(9)
   RTMod.SelStart = 0
   cmdFindNext_Click
   RTModColor = 0
End Sub

Private Sub LIST_PROCS_MODCOLLECTION()
' Called from:-
' mnuAllModStats
' cmdCollect_Click
Dim A$, B$
Dim p As Long     ' General position
Dim k As Long     ' ProcName$(k) number
Dim pSOL As Long  ' Ptr Start of Line
Dim pEOL As Long  ' Ptr End of Line
Dim ModSize As Long
Dim NM As Long    ' Module number
   ReDim NumProcs(ProcSub3)
   ListProcs.Clear
   For NM = 1 To NumMods
      If NM > 1 Then ListProcs.AddItem ""
      ListProcs.AddItem String$(12, "<") & " " & ModName$(NM) & " " & String$(50, ">") '''''''''''''''
      'ListProcs.AddItem ModName$(NM)
      ListProcs.AddItem ""
      ' Extract module string from collection
      ' & calc StartOfProcsPos
      EXTRACT_MODSTRING ModName$(NM)
      ModSize = Len(ModString$)
      ' READ line at a time from ModString$ into A$
      ' & display declarations in ListProcs
      ' Find Declarations
      If LenB(ModString$) > 0 Then
         pSOL = 1
         If StartOfProcsPos > pSOL Then
            Do
               pEOL = InStr(pSOL, ModString$, Chr$(10))
               If pEOL - pSOL - 1 < 1 Then Exit Do
               A$ = Mid$(ModString$, pSOL, pEOL - pSOL - 1)   ' No crlf
               For k = PubS2 To ProcSub1  ' Public Declare Sub -- Public,Private
                  If ProcName$(k) <> "" Then  ' Break
                     p = InStr(1, A$, ProcName$(k))
                     If p = 1 Then
                        NumProcs(k) = NumProcs(k) + 1
                        ListProcs.AddItem A$  '''''''''''''
                        Exit For
                     End If
                  End If
               Next k
               pSOL = pEOL + 1
               If pSOL >= ModSize Then Exit Do
               If pSOL >= StartOfProcsPos Then Exit Do
            Loop
         End If
         ' List Procs
         If StartOfProcsPos > 0 Then
            pSOL = StartOfProcsPos
            If pSOL < ModSize Then      ' To allow for no procs
               Do
                  pEOL = InStr(pSOL, ModString$, Chr$(10))
                  If pEOL - pSOL - 1 < 1 Then Exit Do
                  A$ = Mid$(ModString$, pSOL, pEOL - pSOL - 1)   ' No crlf
                  For k = 1 To PriE1  ' Public Sub -- Friend Property
                     p = InStr(1, A$, ProcName$(k))
                     If p = 1 Then
                        NumProcs(k) = NumProcs(k) + 1
                        ListProcs.AddItem A$  '''''''''''''
                        Exit For
                     End If
                  Next k
                  pSOL = pEOL + 1
                  If pSOL >= ModSize Then Exit Do
               Loop
            End If
         End If
      End If
   Next NM
End Sub

'#### UNUSED DIMS & ARGS #########################################################

Private Sub mnuUNUB_Click(Index As Integer)
' Unused B
Dim A$
Dim NM As Long
Dim NProcs As Long
Dim TotNumDims As Long
Dim TOTALUNUSED As Long

   If Not aColoringDone Then Exit Sub
   Screen.MousePointer = vbHourglass
   DoEvents

   mnuSaveRTB.Enabled = True
   mnuPrintRTB.Enabled = True
   cmdFind.Enabled = True
   cmdFindNext.Enabled = True
   ListProcs.Clear
   For NM = 1 To NumMods
     ListProcs.AddItem ModName$(NM)
   Next NM
   RTMod.Text = ""
   Select Case Index
   Case 0
      A$ = "UNUSED PROC DIMS"
      RTMod.Text = A$ & vbNewLine
      aBar = False
      LIST_DRCs_MODCOLLECTION 1, NProcs, TotNumDims, TOTALUNUSED
      Label3 = Str$(NProcs) & " Procs with " & Str$(TOTALUNUSED) & " Unused Dims "
   Case 1
      A$ = "UNUSED PROC ARGUMENTS"
      RTMod.Text = A$ & vbNewLine
      aBar = False
      LIST_ARGS_MODCOLLECTION 0, NProcs, TOTALUNUSED
      Label3 = Str$(NProcs) & " Procs with " & Str$(TOTALUNUSED) & " Unused Args  ALL "
   Case 2
      A$ = "UNUSED PROC ARGUMENTS (BAR Button, Shift)"
      RTMod.Text = A$ & vbNewLine
      aBar = True
      LIST_ARGS_MODCOLLECTION 0, NProcs, TOTALUNUSED
      Label3 = Str$(NProcs) & " Procs with " & Str$(TOTALUNUSED) & " Unused Args  BAR "
   Case 3
      mnuListProcs_Details_Click 6
      Exit Sub
   End Select
   RTModState = 1
   RTMod.Text = RTMod.Text & RT$ '''''''''''''''''''
   ColorHeader A$
   RTMod.Refresh
   RTBSizeLines
   RT$ = ""
   RTModState = 1 ' Module or List
   
   AddScroll ListProcs
   Screen.MousePointer = vbDefault
   DoEvents
End Sub
'#### END DIMS, REDIMS & CONSTS ################################################


'#### LISTERS, UNUSED PUBLIC/PRIVATES ####################################################

Private Sub mnuList_Click(Index As Integer)
' List All A
Dim KeyWord$
Dim NEnumTypeStarts As Long
Dim NumInMod As Long
Dim Lab3$
   
   If Not aColoringDone Then Exit Sub
   'NB menu Caption names match the ProcName$()
   '   apart from a trailing space & VARS, but
   '   the menu indexes are different
   
   mnuSaveRTB.Enabled = True
   mnuPrintRTB.Enabled = True
   cmdFind.Enabled = True
   cmdFindNext.Enabled = True
   
   RTMod.Text = ""
   
   KeyWord$ = mnuList(Index).Caption
   KeyWord$ = KeyWord$ & " "
   If InStr(1, KeyWord$, "VARS") > 0 Then
      KeyWord$ = Left$(KeyWord$, InStr(1, KeyWord$, " "))
   End If
   Select Case Index ' mnu filter numbers
   Case 1 To PriE1, PubS2 To PriE2, ProcSub1 - 1, ProcSub1
      ListProcs.Clear
      LabVarProc = " Project Files "
      LIST_PUBPRIV_VARS 2, KeyWord$, NumPubPriVars, NEnumTypeStarts, Lab3$
      RTMod.Text = RT$
      RT$ = ""
      If InStr(1, KeyWord$, "Type ") > 0 Or InStr(1, KeyWord$, "Enum ") > 0 Then
         Label3 = Str$(NumPubPriVars - NEnumTypeStarts) & "  " & Lab3$ & "(s) "
      Else
         Label3 = Str$(NumPubPriVars) & "  " & Lab3$ & "(s) "
      End If
   Case Else   ' Menu breaks
   End Select
   
   ColorHeader SearchText$
   
   RTBSizeLines
   RTModState = 1 ' Module or List
End Sub

Private Sub mnuUNUA_Click(Index As Integer)
' Unused A
Dim A$
Dim UCount As Long
Dim Lab3$
Dim KeyWord$
Dim NEnumTypeStarts As Long
Dim NumInMod As Long

   If Not aColoringDone Then Exit Sub
   'NB menu Caption names match the ProcName$()
   '   apart from a trailing space & VARS, but
   '   the menu indexes are different
   
   If NumMods = 0 Or LenB(FileSpec$) = 0 Then Exit Sub
   
   mnuSaveRTB.Enabled = True
   mnuPrintRTB.Enabled = True
   cmdFind.Enabled = True
   cmdFindNext.Enabled = True
   Screen.MousePointer = vbHourglass
   DoEvents
   KeyWord$ = mnuUNUA(Index).Caption
   KeyWord$ = KeyWord$ & " "
   If InStr(1, KeyWord$, "VARS") > 0 Then
      KeyWord$ = Left$(KeyWord$, InStr(1, KeyWord$, " "))
   End If
   ' Index 24 = a break
   Select Case Index
   'Case 1 To 3, 10 To 16, 25     ' mnu Public filter numbers
   Case PubS1 To PubE1, PubS2 To PubE2, ProcSub1 - 1  ' mnu Public/Static filter numbers
      ListProcs.Clear
      LabVarProc = " Project Files "
      LIST_PUBPRIV_VARS 1, KeyWord$, NumPubPriVars, NEnumTypeStarts, Lab3$   ' 1 No print to RT$->RTMod
      If LenB(ModCollection$) = 0 Then
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      RTMod.Text = ""
      RTMod.Refresh
      RT$ = ""
      If KeyWord$ = "Public " Or KeyWord$ = "Private " Then
         KeyWord$ = "UNUSED " & UCase$(KeyWord$) & "VARS"
         'RTMod.Text = KeyWord$ & vbNewLine
         RT$ = KeyWord$ & vbNewLine
         A$ = RT$
         Lab3$ = StrConv(KeyWord$, vbProperCase)
      Else
         KeyWord$ = "UNUSED " & UCase$(KeyWord$) & "(S)"
         'RTMod.Text = KeyWord$ & vbNewLine
         RT$ = KeyWord$ & vbNewLine
         A$ = RT$
         Lab3$ = StrConv(KeyWord$, vbProperCase)
      End If
      EXTRACT_UNUSED_PUBLICS UCount
      RTMod.Text = RT$
      RTMod.Refresh
      RT$ = ""
      Label3 = Str$(UCount) & "  " & Lab3$ & " "
   Case PriS1 To PriE1, PriS2 To PriE2, ProcSub1    ' mnu Private filter numbers
      ListProcs.Clear
      LabVarProc = " Project Files "
      LIST_PUBPRIV_VARS 1, KeyWord$, NumPubPriVars, NEnumTypeStarts, Lab3$   ' 1 No print to RT$->RTMod
      If LenB(ModCollection$) = 0 Then
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      RTMod.Text = ""
      RTMod.Refresh
      RT$ = ""
      If KeyWord$ = "Public " Or KeyWord$ = "Private " Then
         KeyWord$ = "UNUSED " & UCase$(KeyWord$) & "VARS"
         RT$ = KeyWord$ & vbNewLine
         A$ = RT$
         Lab3$ = StrConv(KeyWord$, vbProperCase)
      Else
         KeyWord$ = "UNUSED " & UCase$(KeyWord$) & "(S)"
         RT$ = KeyWord$ & vbNewLine
         A$ = RT$
         Lab3$ = StrConv(KeyWord$, vbProperCase)
      End If
      EXTRACT_UNUSED_PRIVATES UCount
      RTMod.Text = RT$
      RTMod.Refresh
      RT$ = ""
      Label3 = Str$(UCount) & "  " & Lab3$ & " "
   Case Else   ' Menu breaks
   End Select
   ColorHeader A$
   RTModState = 1 ' Module or List
   RTBSizeLines
   
   AddScroll ListProcs
   Screen.MousePointer = vbDefault
End Sub

Private Sub LIST_PUBPRIV_VARS(IPrint As Integer, KeyWord$, NumPubPriVars As Long, NEnumTypeStarts As Long, Lab3$)
' Called from:
' mnuList_Click
' mnuUNUA_Click

' Operating on concatenated stripped modules
' IPrint = 0 Print to RTMod
' IPrint = 1 No print to RTMod
' IPrint = 2 No print for Collect but Print list to RTMod
' eg
' KeyWord$ = "Public "
' KeyWord$ = "Private "

' OUTPUTS:
' NumPubPriVars
' ModNameStore$(NumPubPriVars)
' PubPrivStore$(NumPubPriVars)

Dim NM As Long
Dim NumInMod As Long
Dim Filt1A As Long, Filt2A As Long
Dim Filt1B As Long, Filt2B As Long
Dim aFilter As Boolean
Dim ProcSubscript As Long
   
   Screen.MousePointer = vbHourglass
   DoEvents
   Lab3$ = KeyWord$
   If NumMods = 0 Or LenB(FileSpec$) = 0 Then
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   If LenB(ModCollection$) = 0 Then
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   ReDim ModNameStore$(100)
   ReDim PubPrivStore$(100)
   
   For ProcSubscript = 1 To ProcSub3
      If ProcName$(ProcSubscript) = KeyWord$ Then Exit For
   Next ProcSubscript
   
   If IPrint <> 1 Then
      RT$ = ""  ''''''''''''''''''''''''''
   End If
   
   aFilter = True
   If KeyWord$ = "Public " Then    ' Filter out other Publics
      Filt1A = PubS1: Filt2A = PubE1
      Filt1B = PubS2: Filt2B = PubE2
      Lab3$ = KeyWord$ & "VARS"
      If IPrint <> 1 Then RT$ = UCase$(KeyWord$) & "VARS" & vbNewLine  ''''''''''''''''''''''''''
   ElseIf KeyWord$ = "Private " Then    ' Filter out other Privates
      Filt1A = PriS1: Filt2A = PriE1 + 1   ' Includes Break @ 13
      Filt1B = PriS2: Filt2B = PriE2
      Lab3$ = KeyWord$ & "VARS"
      If IPrint <> 1 Then RT$ = UCase$(KeyWord$) & "VARS" & vbNewLine  ''''''''''''''''''''''''''
   Else
      aFilter = False
      If IPrint <> 1 Then RT$ = UCase$(KeyWord$) & "(S)" & vbNewLine  ''''''''''''''''''''''''''
   End If
   SearchText$ = RT$
   
   NumPubPriVars = 0
   For NM = 1 To NumMods
      NumInMod = 0
      EXTRACT_MODSTRING ModName$(NM)
      ' Returns ModString$ & StartOfProcsPos in ModString$
      GET_VARS NM, IPrint, KeyWord$, NumPubPriVars, NEnumTypeStarts, Filt1A, Filt2A, Filt1B, Filt2B, aFilter, ProcSubscript, NumInMod
      If NumInMod > 0 Then
         ListProcs.AddItem ModName$(NM)
      End If
   Next NM  ' For NM = 1 To NumMods
   ' Return with RT$
   Screen.MousePointer = vbDefault
End Sub
'#### END LISTERS ####################################################

Private Sub mnuExit_Click()
   Form_Unload 0
End Sub

Private Sub mnuPrintRTB_Click()
Dim res As Long
Dim APErr As Boolean
Dim PageCount As Long
   If Not aColoringDone Then Exit Sub
   If Len(RTMod.Text) = 0 Then
      MsgBox "Nothing to print", vbInformation, "Printing RTB"
      Exit Sub
   End If
   PageCount = SendMessageLong(RTMod.hwnd, EM_GETLINECOUNT, 0&, 0&) / 66
   res = MsgBox(" Approximately   " & Str$(PageCount + 1) & " pages." & vbCrLf _
         & " or selected lines." & vbCrLf & vbCrLf _
         & " CONTINUE?   IS PRINTER LIVE!   ", _
      vbQuestion + vbYesNo + vbSystemModal + vbDefaultButton2, "CodeScan - Printing")
   If res = vbYes Then
      ShowPrinter Me, APErr
      If Not APErr Then
         Printer.Print ""
         RTMod.SelPrint Printer.hDC, False  ' False prevents StartDoc & EndDoc
         Printer.EndDoc
      End If
      DoEvents
   End If
End Sub

Private Sub mnuSaveRTB_Click()
Dim Title$, Filt$, InDir$
Dim Ext$
Dim FIndex As Long
Dim p As Long
Dim res As Long
   If Not aColoringDone Then Exit Sub
Set CommonDialog1 = New OSDialog
   Title$ = "Save As *.txt"
   Filt$ = "Save(*.txt)|*.*"
   If FileSpec$ = "" Then
      InDir$ = PathSpec$
   Else
      FixExtension FileSpec$, ".txt"
      p = InStrRev(FileSpec$, "\")
      InDir$ = Left$(FileSpec$, p)
   End If
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
Set CommonDialog1 = Nothing
   If LenB(FileSpec$) > 0 Then
      FixExtension FileSpec$, ".txt"
      RTMod.SaveFile FileSpec$, rtfText
   End If
End Sub

' FONTS
Private Sub cboFS_Click()
' Public GenFontSize
Dim i As Long
   If Not aColoringDone Then
      Exit Sub
   End If
   GenFontSize = Val(cboFS.Text)
   RTMod.Font.Size = GenFontSize
   ListMods.Font.Size = GenFontSize
   ListProcs.Font.Size = GenFontSize
   
   RTBLN.Font.Size = GenFontSize

   RefreshLists
End Sub

Private Sub chkRTB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
' Public aGenBold
   If Not aColoringDone Then
      chkRTB(Index).Value = 1 - chkRTB(Index).Value
      Exit Sub
   End If
   If Index = 0 Then
      aGenBold = -chkRTB(0)
      RTMod.Font.Bold = aGenBold
      ListMods.Font.Bold = aGenBold
      ListProcs.Font.Bold = aGenBold
      RefreshLists
   Else
      RTMod.Font.Underline = -chkRTB(1)
      ListMods.Font.Underline = -chkRTB(1)
      ListProcs.Font.Underline = -chkRTB(1)
   End If
   '   On Error Resume Next

   RTMod.SetFocus
End Sub

Private Sub RefreshLists()
   ListMods.Refresh
   ListProcs.Refresh
   DoEvents
   AddScroll ListMods
   AddScroll ListProcs
   DoEvents
   ListMods.Refresh
   ListProcs.Refresh
   '   On Error Resume Next

   RTMod.SetFocus
   DoEvents
End Sub
'####  RESIZING ######################################################

Private Sub Form_Resize()
   If Not aColoringDone Then Exit Sub
   ResizeControls
   AddScroll ListProcs
End Sub

''-----------------------------------------------
'' Resizing code modification of that by Justin Manley A1VBCODE

' Declare:
'Private Type CScales
'  zLef As Single
'  zTop As Single
'  zWid As Single
'  zHit As Single
'End Type
'Dim ScaleArray() As CScales
' &
' Call InitResizeArray @ Start of prog
'    Form_Activate or Form_Load
'    provided all the controls are
'    present.  Only resizes main
'    Form & its resizable controls

Private Sub InitResizeArray()
'Exit Sub
Dim i As Long

   On Error Resume Next
   ReDim ScaleArray(0 To Controls.Count - 1)

   For i = 0 To Controls.Count - 1
      With ScaleArray(i)
         .zLef = Controls(i).Left / ScaleWidth
         .zTop = Controls(i).Top / ScaleHeight
         .zWid = Controls(i).Width / ScaleWidth
         .zHit = Controls(i).Height / ScaleHeight
         .zFSize = 8 / ScaleWidth
      End With
   Next i
   zRTBFS = ScaleArray(1).zFSize
End Sub

Private Sub ResizeControls()
'Exit Sub
Dim i As Long
'Dim zRTBFS As Single

On Error Resume Next ' Required for controls that wont .Move ,,,

   'Optional: lo-limit form size
   If Form1.Width < ORGFrmWid Or Form1.Height < ORGFrmHit Then
      Form1.Width = ORGFrmWid
      Form1.Height = ORGFrmHit
      Exit Sub
   End If

   Unload frmStats

   '---------------------------------------
   For i = 0 To Controls.Count - 1
      With ScaleArray(i)
         If TypeOf Controls(i) Is ComboBox Or _
            TypeOf Controls(i) Is CheckBox Then
            ' Resize the combo box window with API
            MoveWindow Controls(i).hwnd, .zLef * ScaleWidth, .zTop * ScaleHeight, _
             .zWid * ScaleWidth, .zHit * ScaleHeight, 1
         Else
            If Controls(i).Name <> "cmdMaxRTB" Then
               Controls(i).Move .zLef * ScaleWidth, .zTop * ScaleHeight, _
               .zWid * ScaleWidth, .zHit * ScaleHeight
            End If
         End If
         
         
         Dim zFS As Single
         If TypeOf Controls(i) Is CommandButton Or _
         TypeOf Controls(i) Is Label Or _
         TypeOf Controls(i) Is CheckBox Or _
         TypeOf Controls(i) Is ComboBox Or _
         TypeOf Controls(i) Is Form Then
               zFS = .zFSize * Sqr(ScaleWidth * ScaleHeight)
               If zFS < 8 Then zFS = 8
               If zFS > 12 Then zFS = 12
               Controls(i).FontSize = CInt(zFS)
               If ScaleWidth > 1.5 * (ORGFrmWid / STX) Then
                  Controls(i).FontBold = True
               Else
                  Controls(i).FontBold = False
               End If
         End If
      End With
   Next i
   ' Keep same size to top-right of RTMod
   cmdMaxRTB.Top = RTMod.Top - cmdMaxRTB.Height - 2
   cmdMaxRTB.Left = RTMod.Left + RTMod.Width - cmdMaxRTB.Width - 2
   
' Optional
   With RTBLN
      .Top = RTMod.Top
      .Width = 64
      .Height = RTMod.Height
      .Left = RTMod.Left - RTBLN.Width - 1
   End With
   On Error GoTo 0
End Sub


'### ADDITIONS ##########################################################

Private Sub Timer1_Timer()
'Public FirstVisLine As Long
'Public FoundLine As Long
'Public PrevFirstVisLine As Long
Dim N$
Dim B$
Dim i
   ' To avoid partial line scroll position but doesn't work on
   ' last down-button at End Of RTB when under horz-scrollbar.
   SendMessageLong RTMod.hwnd, EM_LINESCROLL, 0&, ByVal 0
   FirstVisLine = SendMessageLong(RTMod.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
   If FirstVisLine <> PrevFirstVisLine Then
'      SendMessageLong RTMod.hWnd, EM_SCROLL, 0&, ByVal 0
      FirstVisLine = SendMessageLong(RTMod.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
      N$ = ""
      For i = FirstVisLine To FirstVisLine + 100
         B$ = Str$(i + 1)
         If Len(B$) < 6 Then B$ = String$(6 - Len(B$), " ") & B$
         N$ = N$ & Format(B$, "@@@@@") & vbCrLf
      Next i
      RTBLN.Text = N$
      PrevFirstVisLine = FirstVisLine
   End If
End Sub

Private Sub chkLineNumbers_Click()
   If Not aColoringDone Then
      chkLineNumbers.Value = 0
      Exit Sub
   End If
   'aTimer = -chkLineNumbers.Value
   Timer1.Enabled = Not Timer1.Enabled
   If Not Timer1.Enabled Then
      aTimer = False
      RTBLN.Text = ""
   Else
      If Len(RTMod.Text) > 0 Then
         aTimer = True
         PrevFirstVisLine = -1
         Timer1_Timer
      Else
         Timer1.Enabled = False
         aTimer = False
         RTBLN.Text = ""
         chkLineNumbers.Value = 0
      End If
   End If
   On Error Resume Next
   RTMod.SetFocus
End Sub

Private Sub cmdColorIt_Click()
Dim A As Boolean
   'If Not aColoringDone Then Exit Sub
   A = aTimer
   Screen.MousePointer = vbHourglass
   If aTimer Then
      aTimer = False
      chkLineNumbers.Value = 0
   End If
   
   ColorRTB Form1, RTMod
   aColoringDone = True
   
   If A Then
      aTimer = True
      chkLineNumbers.Value = 1
   End If
   RTMod.SetFocus
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdStopColoring_Click()
   aColoringDone = True
End Sub

'Private Sub cmdHSSBS_Click(Index As Integer)
Private Sub cmdHSSBS_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Not aColoringDone Then Exit Sub
   If Index = 0 Then
      ListBoxHorScroll ListMods, False
      ListProcs.SetFocus
      ListBoxHorScroll ListProcs, False
      RTBHorScroll RTMod, False
   Else
      ListBoxHorScroll ListMods, True
      ListProcs.SetFocus
      ListBoxHorScroll ListProcs, True
      RTBHorScroll RTMod, True
   End If
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
      SendMessage RTMod.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, RTFformat
      RTMod.SelStart = FoundPos   ' set selection start
   End If
End Sub

'########################################################################

Private Sub Form_Unload(Cancel As Integer)
Dim Form As Form
   aColoringDone = True
   ChDir PathSpec$
   Open PathSpec$ & "CSInfo.txt" For Output As #1
   Print #1, FileSpec$
   Close
   ModCollection$ = ""
   ModString$ = ""
   RT$ = ""
   DoEvents
   ' Make sure all forms cleared
   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form
   End
End Sub

