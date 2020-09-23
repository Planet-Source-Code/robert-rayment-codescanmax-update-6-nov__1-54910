VERSION 5.00
Begin VB.Form frmStats 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmStats"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   4230
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   577
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   282
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmStats.frm

Option Explicit

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Const hWndInsertAfter = -1 ' On top
'Private Const wFlags = &H1 Or &H2 Or &H40  ' No move & Show window
Private Const wFlags = &H1 Or &H40   ' No move & Show window

Dim aa As Long
Dim bb As Long

Private Sub Form_Load()
' Public
' ReDim ModuleType$(4)
' ReDim NumModTypes(4)
' ReDim ProcName$(ProcSub3)
' ReDim NumProcs(ProcSub3) As Long
Dim k As Long
Dim p1 As Long
Dim p2 As Long

   ' Public
   'frmStats.Width = 2655
   'frmStats.Height = 8310
   frmStats.Width = frmStatsWidth
   frmStats.Height = frmStatsHeight
   
   
   If Form1.WindowState <> vbMaximized Then
      frmStats.Left = Form1.Left - frmStatsWidth
      frmStats.Top = Form1.Top
   Else  ' Maximised  NB frmStats travels when repeated Unload/Load ?
      'frmStatsLeft = 40
      'frmStatsTop = 66
      ' Size & Make frmStats stay on top
      SetWindowPos frmStats.hwnd, hWndInsertAfter, frmStatsLeft / STX, frmStatsTop / STY, _
      frmStatsWidth / STX, frmStatsHeight / STY, wFlags
   End If
   
'TEST
'Form1.Label3 = Str$(Form1.Left) & Str$(Form1.Top)
   
   Caption = "Stats"
   Cls
   ForeColor = vbRed
   FontBold = True
   Print GetFileName(FileSpec$)
   FontBold = False
   Print
   For k = 1 To NModTypes
      Print ModuleType$(k); Tab(26); NumModTypes(k)
   Next k
   Print Tab(26); "___"
   Print Tab(26); NumMods
   
   If ModFileSpec$ = "" Then Exit Sub
   ForeColor = vbBlack
   FontBold = True
   Print GetFileName(ModFileSpec$)
   FontBold = False
   Print
   For k = 1 To ProcSub1  'ProcSub3 - 2
      If ProcName$(k) <> "" Then ' Menu Breaks
         p1 = InStr(1, ProcName$(k), "Enum ")
         p2 = InStr(1, ProcName$(k), "Type ")
         If p1 = 0 And p2 = 0 Then
            If k = ProcSub1 - 1 Or k = ProcSub1 Then
               Print ProcName$(k); " VARS"; Tab(26); NumProcs(k)
            Else
               Print ProcName$(k); Tab(26); NumProcs(k)
            End If
         Else
            Print ProcName$(k) + " (starts)"; Tab(26); NumProcs(k)
         End If
         If k = PriE1 Then Print
      End If
   Next k
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmStatsWidth = frmStats.Width
   frmStatsHeight = frmStats.Height
   frmStatsLeft = frmStats.Left
   frmStatsTop = frmStats.Top
   Unload frmStats
End Sub
