Attribute VB_Name = "Lister"
' Lister.bas

Option Explicit
Option Base 1
Private NUM_PROCS_USED_IN_MOD As Long
Private ALL_NUM_PROCS_USED As Long
Private DRCStore$()
Private DRCType() As Long

Public Sub UnusedProc()

End Sub
Public Sub GET_MODSTATS(ModuleName$)
' Called from:-
' mnuModStats
Dim A$
Dim p As Long
Dim k As Long
Dim j1 As Long
Dim j2 As Long
Dim NM As Long
   ReDim NumProcs(ProcSub3)
   
   A$ = ModuleName$
   EXTRACT_MODSTRING A$
   ' Returns ModString$ & StartOfProcsPos in ModString$
   
   If LenB(A$) = 0 Then
      MsgBox ModuleName$ & " Select a module ", vbCritical, " Get module name"
      Exit Sub
   End If
         
   ' READ line at a time from ModString$ into A$
   ' & count NumProcs(k) in module
   If LenB(ModString$) > 0 Then
      j1 = 1
      Do
         j2 = InStr(j1, ModString$, Chr$(10))
         If j2 - j1 - 1 < 1 Then Exit Do
         A$ = Mid$(ModString$, j1, j2 - j1 - 1)
         For k = 1 To ProcSub1 ' Public Declare Sub -- Public,Private
            If ProcName$(k) <> "" Then ' Breaks
               If InStr(1, A$, ProcName$(k)) = 1 Then
                  NumProcs(k) = NumProcs(k) + 1
                  Exit For
               End If
            End If
         Next k
         j1 = j2 + 1 ' Start of next line
         If j1 >= Len(ModString$) Then Exit Do
      Loop
   End If
End Sub

Public Sub EXTRACT_UNUSED_PUBLICS(UCount As Long)
' Called from:
' mnuUNUA_Click

Dim k As Long
Dim j1 As Long

On Error GoTo PUBLIC_EXTRACT_ERR

   UCount = 0
   ModString$ = ModCollection$
   For k = 1 To NumPubPriVars
      'SORT uses:- ModString$, PubPrivStore$(k)
      SORT UCount, k
   Next k
   
   Screen.MousePointer = vbDefault
   Exit Sub
'===========
PUBLIC_EXTRACT_ERR:
MsgBox "Unused public extractor error", vbCritical, "CodeScan"
Resume Next
End Sub

Public Sub EXTRACT_MODSTRING(ModuleName$)
' Extracts stripped module ModuleName$ from ModCollection$
' OUT: Public ModString$ + StartOfProcsPos
' Called from:
' GET_MODSTATS
' LIST_PUBPRIV_VARS
' EXTRACT_UNUSED_PRIVATES
' cmdStrip_Click
' LIST_DRCs_MODCOLLECTION

Dim NM As Long
   'ReDim NumProcs(ProcSub3)
   For NM = 1 To NumMods
      If ModName$(NM) = ModuleName$ Then
         ' Extract module string from collection
         ' & calc StartOfProcsPos
         If NM < NumMods Then
            If ModStartPos(NM) = 0 Or ModStartPos(NM + 1) = 0 Then
               ModString$ = ""
               StartOfProcsPos = 0
            Else
               ModString$ = Mid$(ModCollection$, ModStartPos(NM), ModStartPos(NM + 1) - ModStartPos(NM))
               StartOfProcsPos = ModProcPos(NM) - ModStartPos(NM) + 1
            End If
         Else
            If ModStartPos(NM) = 0 Then
               ModString$ = ""
               StartOfProcsPos = 0
            Else
               ModString$ = Mid$(ModCollection$, ModStartPos(NM), Len(ModCollection$) - ModStartPos(NM) + 1)
               StartOfProcsPos = ModProcPos(NM) - ModStartPos(NM) + 1
            End If
         End If
         Exit For
      End If
   Next NM
   If NM = NumMods + 1 Then ModuleName$ = ""
   
End Sub

Public Sub EXTRACT_UNUSED_PRIVATES(UCount As Long)
' Called from:
' mnuUNUA_Click

' Enter with Public:
' NumMods
' NumPubPriVars
' ModNameStore$(NumPubPriVars) = ModName$(NM)
' PubPrivStore$(NumPubPriVars)

' UCount - unused counter
Dim k As Long
Dim NM As Long

On Error GoTo PRIVATE_EXTRACT_ERR

   UCount = 0
   For NM = 1 To NumMods
      If NumPubPriVars > 0 Then
         EXTRACT_MODSTRING ModName$(NM)
         ' Returns Stripped module in ModString$ & StartOfProcsPos
         
         If LenB(ModString$) > 0 Then
            For k = 1 To NumPubPriVars
               If ModName$(NM) = ModNameStore$(k) Then
                  'SORT uses:- ModString$, PubPrivStore$(k)
                  SORT UCount, k
               End If
            Next k
         End If
      End If
   Next NM
   Screen.MousePointer = vbDefault
   Exit Sub
'===========
PRIVATE_EXTRACT_ERR:
MsgBox "Unused private extractor error", vbCritical, "CodeScan"
Resume Next
End Sub

Public Sub SORT(UCount As Long, k As Long)
' Called from:
' EXTRACT_UNUSED_PUBLICS
' EXTRACT_UNUSED_PRIVATES

' Determines whether or not PubPrivStore$(k) exists
' more than once in ModString$

' Input: Public ModString$
'        UCount = unused count
'        k for ModNameStore$(k) & PubPrivStore$(k)
' Out:   UCount

Dim Bef$, First$, Last$, Aft$
Dim p As Long
Dim UseCount As Long
Dim aKey As Boolean
Dim A$, B$
Dim aInsertSpace As Boolean
   
    aInsertSpace = False
   
   ' READ line at a time from ModString$ into A$
   ' & display declarations in ListProcs
   ' Find prelims
   UseCount = 0
   p = 0
   First$ = Left$(PubPrivStore$(k), 1)
   'Last$ = Right$(PubPrivStore$(k), 1)
   If First$ <> "." Then      'eg Type            if Type used var needed for
                              '     var As Long    structure even if not used,
                              '                    .var saved in PubPrivStore$(k)
      If First$ = "*" Then    ' Enum marker
         aInsertSpace = True
         PubPrivStore$(k) = Mid$(PubPrivStore$(k), 2) ' Remove * flag
      End If
      
      Do
         p = InStr(p + 1, ModString$, PubPrivStore$(k))
         If p > 0 Then
            If p > 1 Then
            
               Bef$ = Mid$(ModString$, p - 1, 1)
               Aft$ = Mid$(ModString$, p + Len(PubPrivStore$(k)), 1)
               Last$ = Mid$(ModString$, p + Len(PubPrivStore$(k)) - 1, 1)
               
               aKey = TEST_CHAR(Bef$, Aft$, Last$)
               
               If aKey Then
                  UseCount = UseCount + 1
               End If
            
            End If
         Else
            Exit Do
         End If
         
         If UseCount > 1 Then Exit Do  ' ie has been used more than once
      Loop Until p <= 0
   
   End If
   If UseCount = 1 Then
      A$ = PubPrivStore$(k)
      If aInsertSpace Then
         A$ = " " & A$
         aInsertSpace = False
      End If
      B$ = ModNameStore$(k)
      If Len(B$) > 28 Then
         B$ = Left$(B$, 11) & ".." & Right$(B$, 15)
      End If
      
      'RTMod.Text = RTMod.Text & B$ & Space$(30 - Len(B$)) & A$ & vbNewLine '-------------------
      
      RT$ = RT$ & B$ & Space$(30 - Len(B$)) & A$ & vbNewLine '-------------------
      
      UCount = UCount + 1
   End If
End Sub

Public Sub LIST_CTRL_NAMES(NCtrls As Long)
Dim A$
Dim k As Long
Dim NM As Long
Dim NCtrl As Long

   RT$ = ""
   NCtrl = UBound(ModCtrlName$(), 2)
   NCtrls = 0
   For NM = 1 To NumMods
      RT$ = RT$ & ModName$(NM) & vbNewLine
      For k = 1 To NCtrl
         A$ = ModCtrlName$(NM, k)
         If A$ <> "" Then
            NCtrls = NCtrls + 1
            RT$ = RT$ & Space$(12) & A$ & vbNewLine
         End If
      Next k
   Next NM
   ' Return with RT$
End Sub

Public Sub EXTRACT_DRC_NAMES(B$, C$)
' IN:  B$ string to check
' OUT: C$ Name only
' Check order:
' 1 Dim a(.,.)
' 2 Dim a As Long
' 3 Dim a$crlf
'   ReDim vv(
'   Const xx =
Dim p As Long
   C$ = ""
   p = InStr(1, B$, " ")
   If p = 0 Then Exit Sub     'ERROR
   C$ = Mid$(B$, p + 1)
   p = InStr(1, C$, "(")
   If p > 0 Then
      C$ = Left$(C$, p)
   Else
      p = InStr(1, C$, " ")
      If p <> 0 Then
         C$ = Left$(C$, p - 1)
      Else
         p = InStr(1, C$, Chr$(13))
         If p = 0 Then Exit Sub   'ERROR
         C$ = Left$(C$, p - 1)
      End If
  End If
End Sub

'Public Sub FIND_UNUSED_DIMS(PString$, DStore$(), NDimsInProc As Long, UNUSED As Long)
Public Sub FIND_UNUSED_DIMS(PString$, NDimsInProc As Long, UNUSED As Long)
' Private DRCStore$()

' Called from:
' LIST_DRCs_MODCOLLECTION

'                Input                    Input/Output   Input          OutPut
' FIND_UNUSED_DIMS Chr$(10) & ProcString$,  DRCStore$(),   NumDRCsInProc, UNUSED

ReDim Dummy$(NDimsInProc) ' To collect unused Dim Vars
Dim NU As Long
Dim A$
Dim aKey As Boolean
Dim k As Long
Dim kk As Long
Dim p As Long
Dim p1 As Long
Dim pSOL As Long  ' Ptr Start of Line
Dim UseCount As Long
Dim pInLine As Long  ' Ptr In Line
Dim Bef$
Dim Aft$
Dim Last$

NU = 0
Last$ = ""

   For k = 1 To NDimsInProc
         'A$ = DStore$(k)
         A$ = DRCStore$(k)
         pSOL = 1
         UseCount = 0
         Do
            ' Extract a line from PString$
            p = InStr(pSOL, PString$, Chr$(13))
            A$ = Mid$(PString$, pSOL, p - pSOL)
   
               pInLine = 1
               ' See if DStore$(k) is in this line
               Do
                  p = InStr(pInLine, A$, DRCStore$(k))
                  Last$ = Right$(DRCStore$(k), 1)
                  
                  If p > 0 Then
                     
                     Bef$ = Mid$(PString$, pSOL + p - 2, 1)
                     Aft$ = Mid$(PString$, pSOL + p - 1 + Len(DRCStore$(k)), 1)
                     
                     aKey = TEST_CHAR(Bef$, Aft$, Last$)
                     
                     If aKey Then   ' Count uses of DStore$(k)
                        UseCount = UseCount + 1
                     End If
                  Else
                     If Last$ = "(" Then     ' eg Dim Array() used as Array = or (Array)
                        p1 = InStr(pInLine, A$, Left$(DRCStore$(k), Len(DRCStore$(k)) - 1))
                        If p1 > 0 Then
                           Last$ = ""
                           Bef$ = Mid$(PString$, pSOL + p1 - 2, 1)
                           Aft$ = Mid$(PString$, pSOL + p1 - 1 + Len(DRCStore$(k)) - 1, 1)
                           
                           aKey = TEST_CHAR(Bef$, Aft$, Last$)
                           
                           If aKey Then   ' Count uses of DStore$(k)
                              UseCount = UseCount + 1
                           End If
                        Else
                           Exit Do  ' No DStore$(k) in line
                        End If
                     End If
                  End If
                  
                  If UseCount > 1 Then Exit Do  ' ie has been used more than once
                  ' Move along line for any more uses of DStore$(k)
                  pInLine = p + Len(DRCStore$(k))
                  If pInLine >= Len(A$) Then Exit Do
               Loop Until p <= 0
            
            If UseCount > 1 Then
               Exit Do     ' DRCStore$(k) used test next DRCStore$(k+1)
            End If
            ' Get next line in PString$ & see if DStore$(k) used in that
            pSOL = pSOL + Len(A$) + 2
            If pSOL >= Len(PString$) Then Exit Do
         Loop
   
         If UseCount = 1 Then ' DStore$(k) var only used once store it
            NU = NU + 1
            Dummy$(NU) = DRCStore$(k)
         End If
   ' Get next DStore$(k)
   Next k
   
   If NU > 0 Then
      If NU > UBound(DRCStore$(), 1) Then
         ReDim Preserve DRCStore$(NU + 20)
      End If
      For k = 1 To NU
         DRCStore$(k) = Dummy$(k)
      Next k
   End If
   Erase Dummy$()
   UNUSED = NU ' >= 0
      
End Sub

Public Function TEST_CHAR(Bef$, Aft$, Last$) As Boolean
' Called from:
' SORT
' FIND_UNUSED_DIMS
' FIND_UNUSED_ARGS

Dim aKey As Boolean
   
   ' Oh Boy!!
   ' Test either side of possible var string
   ' ^=Space
   ' Before    After
   ' ^         ^ , ( ) ; . CR
   ' LF        ^ ( . CR
   ' (         ) ( , ^ . :
   ' -         ^
   ' .         ^
   
   
   aKey = False
   If Last$ = "(" Then  ' Since arg might be ^var( and use ^var(k) so Aft$=k OR var no (
      Select Case Bef$
      Case " ", "(", Chr$(10)  'eg xx = yy(8) | yy = xx(yy(8)) |lf y(8) = xx
                               '       B  (            B  (      B  (
         aKey = True
      Case Else
         aKey = False
      End Select
   Else
      Select Case Bef$
      Case " "
         Select Case Aft$
         Case " ", ",", "(", ")", ";", ".", Chr$(13)  ' ; eg Print var;
            aKey = True
         Case Else
            aKey = False
         End Select
      Case Chr$(10)
         Select Case Aft$
         Case " ", "(", ".", Chr$(13)
            aKey = True
         Case Else
            aKey = False
         End Select
      Case "("
         Select Case Aft$
         Case ")", "(", ",", " ", ".", ":"   ' (NamedArgument:=
            aKey = True
         Case Else
            aKey = False
         End Select
      Case "-"
         Select Case Aft$
         Case " "
            aKey = True
         Case Else
            aKey = False
         End Select
      Case "."
         Select Case Aft$
         Case " ", ".", "(", Chr$(13)
            aKey = True
         Case Else
            aKey = False
         End Select
      End Select
   End If
   TEST_CHAR = aKey
End Function

Public Sub LIST_ARGS_MODCOLLECTION(LISTTYPE As Long, NProcs As Long, TOTALUNUSED As Long)
' LISTTYPE = 0 List Unused
' LISTTYPE = 1 List Procs & Arguments
' Called from:
' mnuListProcs_Dims_Click
' mnuUNUB_Click

Dim A$, B$, C$
Dim p As Long
Dim p2 As Long
Dim p3 As Long
Dim i As Long
Dim j As Long
Dim k As Long

' ModString$ positions
Dim pSOProc As Long  ' Ptr Start of Procs
Dim pEOL As Long     ' Ptr End of Line
Dim pSOL As Long     ' Ptr Start of Line
Dim pSONextL As Long ' Ptr Start of Next Line
Dim pEOProc As Long  ' Ptr End of Proc
Dim ModSize As Long  ' = Len(ModString$)

Dim NM As Long
Dim EndString$
Dim UNUSED As Long
Dim ProcString$
Dim ProcStartLine$
Dim FirstIn As Long

Dim UNArgStore$()

'Public TotNumArgs As Long
'Public NumArgsInProc As Long
'Public ArgStore$()
'Public aBar As Boolean ' To Bar Index, Button, Shift etc
'Public RT$

ReDim ArgStore$(50)
ReDim UNArgStore$(50)

   ReDim NumProcs(ProcSub3)
'   ListProcs.Clear
   RT$ = ""
   TOTALUNUSED = 0
   NProcs = 0
   'LabVarProc = " Project Files "
   For NM = 1 To NumMods
      
      ' Extract module string from collection
      ' into ModString$ & calc StartOfProcsPos
      EXTRACT_MODSTRING ModName$(NM)
      
      ' READ line at a time from ModString$ into A$
      ' & display declarations in ListProcs
      
      If LenB(ModString$) > 0 Then
         
         Form1.Label3 = " " & ModName$(NM) & " "
         Form1.Label3.Refresh
         
         RT$ = RT$ & ModName$(NM) & vbNewLine  ''''''''''''''''''''''''
         
         ModSize = Len(ModString$)
         pSOProc = StartOfProcsPos
         If pSOProc < ModSize Then      ' To allow for no procs
            Do
               pEOL = InStr(pSOProc, ModString$, Chr$(10))
               If pEOL - pSOProc - 1 < 1 Then Exit Do
               ' Collect line
               A$ = Mid$(ModString$, pSOProc, pEOL - pSOProc - 1)
               pSOL = pEOL + 1 ' -> Start of line after Proc Start Line
               
               For k = 1 To PriE1 ' Public Sub -- Friend Property
                  If InStr(1, A$, ProcName$(k)) = 1 Then
                     ' PROC FOUND
                     ' Keep Proc start line
                     FirstIn = 0
                     ProcStartLine$ = Space$(10) & A$ & vbNewLine
                     EndString$ = "End Sub"
                     If InStr(1, ProcName$(k), "Function ") > 0 Then
                        EndString$ = "End Function"
                     ElseIf InStr(1, ProcName$(k), "Property ") > 0 Then
                        EndString$ = "End Property"
                     End If
                     ' Get End of Proc
                     pEOProc = InStr(pSOProc, ModString$, Chr$(10) & EndString$)
                     ' pEOProc ->LF End Sub/Function/Property
                     pEOProc = pEOProc + Len(EndString$) + 2      ' -> next LF
                     EXTRACT_PROC_ARGS A$     ' To ArgStore$() NumArgsInProc
                     
                     If NumArgsInProc > 0 Then
                        
                        If LISTTYPE = 1 Then    ' List Procs & Arguments
                           NProcs = NProcs + 1
                           RT$ = RT$ & ProcStartLine$ ''''''''''''''''''
                           ' Print Args
                           For i = 1 To NumArgsInProc
                              RT$ = RT$ & Space$(12) & ArgStore$(i) & vbNewLine ''''''''''''
                           Next i
                        Else  ' LISTTYPE = 0 List Unused
                           Do
                              ' pSOL Start of current line
                              ' pSONextL = Start of next line
                              pSONextL = InStr(pSOL, ModString$, Chr$(10)) + 1
                              If pSONextL <= pSOL Then Exit Do
                              ' B$ = current line
                              B$ = Mid$(ModString$, pSOL, pSONextL - pSOL)  ' Includes CRLF
                              ' End Sub/Function/Property
                              If InStr(1, B$, EndString$) = 1 Then Exit Do    '>>>
                              
                              UNUSED = 0
                              ProcString$ = Mid$(ModString$, pSOProc, pEOProc - pSOProc)  ' + Len(EndString$) + 1)
                              FIND_UNUSED_ARGS Chr$(10) & ProcString$, ArgStore$(), NumArgsInProc, UNArgStore$(), UNUSED
                              If UNUSED > 0 Then
                                 NProcs = NProcs + 1
                                 FirstIn = FirstIn + 1
                                 If FirstIn = 1 Then RT$ = RT$ & ProcStartLine$ '''''''''''''
                                 TOTALUNUSED = TOTALUNUSED + UNUSED
                                 For p = 1 To UNUSED
                                    RT$ = RT$ & Space$(12) & UNArgStore$(p) & vbNewLine ''''''''''''
                                 Next p
                                 Exit Do '>>>
                              End If   ' If UNUSED > 0
                              pSOL = pSONextL  ' pSOL = Start of next line
                              If pSOL >= ModSize Then Exit Do         ' ModSize = Len(ModString$)
                           Loop
                        End If
                     End If   ' If NumArgsInProc > 0
                     Exit For '>>>
                  End If   ' If InStr(1, A$, ProcName$(k)) = 1
               Next k   ' k = 1 To PriE1 ' Public Sub -- Friend Property
               If k = PriE1 + 1 Then Exit Do
               pSOProc = pEOProc + 1  'Len(EndString$) + 3 'pEOL + 1
               If pSOProc >= ModSize Then Exit Do      '>>>  ModSize = Len(ModString$)
            Loop  ' Loop thru whole ModString$ with next Proc Start line
         End If   ' If pSOProc < ModSize Then
      End If   ' If LenB(ModString$) > 0
   Next NM  ' Next Module
   ' Return with RT$
End Sub

Public Sub EXTRACT_PROC_ARGS(A$)       ' To ArgStore$() NumArgsInProc
' Called from:
' LIST_ARGS_MODCOLLECTION

' IN: a$= Proc start line
Dim p As Long
Dim p2 As Long
Dim B$, C$
Dim i As Long
Dim j As Long
Dim ArgCount As Long

' Public AStore$
   ReDim AStore$(50)
   NumArgsInProc = 0
   p = InStr(1, A$, "(")
   p2 = InStr(1, A$, ")")
   If p2 = p + 1 Then
      NumArgsInProc = 0
   Else
      ' Count Commas not in quotes  ??
      
      ' Replace anything in eg " xxx,  " by "???"
      Dim nq As Long
      Dim cprev$
      nq = 0
      For p2 = p + 1 To Len(A$) - 1
         C$ = Mid$(A$, p2, 1)
         If nq = 1 And cprev$ = Chr$(34) And C$ <> Chr$(34) Then
            Mid$(A$, p2, 1) = "?"
            'Debug.Print A$
         End If
         If C$ = Chr$(34) Then
            nq = nq + 1
            cprev$ = Chr$(34)
            If nq = 2 Then
               nq = 0
               cprev$ = ""
            End If
         End If
      Next p2
      
      For p2 = p + 1 To Len(A$) - 1
         C$ = Mid$(A$, p2, 1)
         If C$ = "," Then
               NumArgsInProc = NumArgsInProc + 1
               If NumArgsInProc > UBound(ArgStore$(), 1) Then
                  ReDim Preserve AStore$(NumArgsInProc)
               End If
               B$ = Mid$(A$, p + 1, p2 - p - 1)
               AStore$(NumArgsInProc) = Mid$(A$, p + 1, p2 - p - 1)
               p = p2 + 1
         End If
      Next p2
      NumArgsInProc = NumArgsInProc + 1
      If NumArgsInProc > UBound(AStore$(), 1) Then
         ReDim Preserve AStore$(NumArgsInProc)
      End If
      ReDim ArgStore$(NumArgsInProc)
      ArgCount = 0
      
      AStore$(NumArgsInProc) = Mid$(A$, p + 1)
      AStore$(NumArgsInProc) = Left$(AStore$(NumArgsInProc), Len(AStore$(NumArgsInProc)) - 1)
      
       ' Check pre-name parameters
      For i = 1 To NumArgsInProc
         p = InStr(1, AStore$(i), " ")
         If p > 0 Then
            ' Check if Optional ByVal/ByRef   ie extra space
            j = InStr(1, AStore$(i), "Optional ByVal ")
            If j > 0 Then
               p = InStr(p + 1, AStore$(i), " ")
            End If
            j = InStr(1, AStore$(i), "Optional ByRef ")
            If j > 0 Then
               p = InStr(p + 1, AStore$(i), " ")
            End If
            
            C$ = Left$(AStore$(i), p - 1)
            If C$ <> "" Then
               Select Case C$
               Case "Optional ByVal", "Optional ByRef"
                  AStore$(i) = Mid$(AStore$(i), p + 1)
                  ' Optional ByVal/ByRef Name As
                  ' Optional ByVal/ByRef Name,
                  ' Optional ByVal/ByRef Name)
                  ' Optional ByVal/ByRef Name(
                  '                     ^p
               Case "Optional", "ByVal", "ByRef", "ParamArray", "AddressOf"
                  '    Optional Name =
                  '  ParamArray Name
                  ' ByVal/ByRef Name As
                  ' ByVal/ByRef Name,
                  ' ByVal/ByRef Name)
                  ' ByVal/ByRef Name(
                  '            ^p
                  AStore$(i) = Mid$(AStore$(i), p + 1)
                  ' Just Name... left
               End Select
            End If
         End If
         ' Just a Name... left
         For j = 1 To Len(AStore$(i))
            C$ = Mid$(AStore$(i), j, 1)
            Select Case C$
            Case " ", ",", ")"
               AStore$(i) = Left$(AStore$(i), j - 1)
               Exit For
            Case "("
               AStore$(i) = Left$(AStore$(i), j)
               Exit For
            Case Else   ' ERROR
            End Select
         Next j
         If aBar Then
            Select Case AStore$(i)
            Case "Button", "Shift"
            'Case "Button", "Shift", "x", "y", "X", "Y"
            Case Else
               ArgCount = ArgCount + 1
               ArgStore$(ArgCount) = AStore$(i)
            End Select
         Else
            ArgCount = ArgCount + 1
            ArgStore$(ArgCount) = AStore$(i)
         End If
      Next i
      NumArgsInProc = ArgCount
   End If
   Erase AStore$()
End Sub

Public Sub FIND_UNUSED_ARGS(PString$, DStore$(), NArgsInProc As Long, UNArgStore$(), UNUSED As Long)
' Called from:
' LIST_ARGS_MODCOLLECTION

'                Input                    Input/Output   Input          Input/OutPut
' FIND_UNUSED_ARGS Chr$(10) & ProcString$, ArgStore$(), NumArgsInProc, UNArgStore$(), UNUSED

ReDim UNArgStore$(NArgsInProc) ' To collect unused Args
Dim NU As Long
Dim A$, B$
Dim aKey As Boolean
Dim k As Long
Dim kk As Long
Dim p As Long
Dim p1 As Long
Dim pSOL As Long  ' Ptr Start of Line
Dim UseCount As Long
Dim pInLine As Long  ' Ptr In Line
Dim Bef$
Dim Aft$
Dim Last$
Dim pos As Long

NU = 0
Last$ = ""

   For k = 1 To NArgsInProc
         
         A$ = DStore$(k)
         pSOL = 1
         UseCount = 0
         Do
            ' Extract a line from PString$
            p = InStr(pSOL, PString$, Chr$(13))
            If p - pSOL < 1 Then Exit Do
            
            B$ = Mid$(PString$, pSOL, p - pSOL)
            pInLine = 1
            Last$ = Right$(DStore$(k), 1)
            ' See if DStore$(k) is in this line
            Do
               p = InStr(pInLine, B$, DStore$(k))
               
               If p > 0 Then
                  pos = pSOL + p - 2
                  If pos = 0 Then pos = 1
                  Bef$ = Mid$(PString$, pos, 1)
                  Aft$ = Mid$(PString$, pSOL + p - 1 + Len(DStore$(k)), 1)
                  
                  aKey = TEST_CHAR(Bef$, Aft$, Last$)
                  
                  If aKey Then   ' Count uses of DStore$(k)
                     UseCount = UseCount + 1
                  End If
               
               Else
                  If Last$ = "(" Then     ' eg ProcName(i, Array() As String)  used as Array = or (Array)
                     p1 = InStr(pInLine, A$, Left$(DStore$(k), Len(DStore$(k)) - 1))
                     If p1 > 0 Then
                        Last$ = ""
                        pos = pSOL + p - 2
                        If pos = 0 Then pos = 1
                        Bef$ = Mid$(PString$, pos, 1)
                        Aft$ = Mid$(PString$, pSOL + p1 - 1 + Len(DStore$(k)) - 1, 1)
                        
                        aKey = TEST_CHAR(Bef$, Aft$, Last$)
                        
                        If aKey Then   ' Count uses of DStore$(k)
                           UseCount = UseCount + 1
                        End If
                     Else
                        Exit Do  ' No DStore$(k) in line
                     End If
                  End If
               End If
               
               If UseCount > 1 Then Exit Do  ' ie has been used more than once
               
               ' Move along line for any more uses of DStore$(k)
               pInLine = p + Len(DStore$(k))
               If pInLine >= Len(B$) Then Exit Do
            Loop Until p <= 0
            
            If UseCount > 1 Then
               Exit Do     ' DStore$(k) used test next DStore$(k+1)
            End If
            
            ' Get next line in PString$ & see if DStore$(k) used in that
            pSOL = pSOL + Len(B$) + 2
            If pSOL >= Len(PString$) Then Exit Do
         Loop
   
         If UseCount = 1 Then ' DStore$(k) var only used once store it
            NU = NU + 1
            UNArgStore$(NU) = DStore$(k)
         End If
   ' Get next DStore$(k)
   Next k
   
   UNUSED = NU ' >= 0

End Sub

Public Sub LIST_PROC_CALLERS(LISTTYPE As Long, NProcs As Long, TOTALUNUSED As Long)
' Called from:
' mnuListProcs_Details_Click
' mnuUNUB_Click

' LISTTYPE = 0   Include Ctrl Procs
' LISTTYPE = 1   Exclude Ctrl Procs ie CtrlName_
' LISTTYPE = 2   List Non-Control Proc Callers

Dim A$, B$
Dim p As Long
Dim p2 As Long
Dim p3 As Long
Dim ps As Long
Dim pe As Long
Dim i As Long
Dim j As Long
Dim k As Long
' ModString$ positions
Dim pSOProc As Long     ' Ptr Start of Proc
Dim pEOL As Long        ' Ptr Start of Line
Dim pSONextL As Long    ' Ptr Start of Next Line
Dim pEOProc As Long     ' Ptr End of Proc
Dim j4 As Long
Dim ModSize As Long

Dim NM As Long
Dim TotNumDims As Long
Dim NumDimsInProc As Long
Dim EndString$
Dim UNUSED As Long
'Dim TOTALUNUSED As Long
Dim ProcString$
Dim ProcStartLine$
Dim FirstIn As Long
Dim pund As Long  ' _ pos
Dim C$     ' Ctrl name
Dim D$
Dim E$     ' Event
Dim TempModString$
Dim TempStartOfProcsPos As Long
   RT$ = ""
   TotNumDims = 0
   TOTALUNUSED = 0
   NProcs = 0
   For NM = 1 To NumMods
      ' Extract module string from collection
      ' into ModString$ & calc StartOfProcsPos
      EXTRACT_MODSTRING ModName$(NM)
      ' READ line at a time from ModString$ into A$
      ' & display declarations in ListProcs
      If LenB(ModString$) > 0 Then
         
         Form1.Label3 = " " & ModName$(NM) & " "
         Form1.Label3.Refresh
         
         If LISTTYPE <> 2 Then
            RT$ = RT$ & ModName$(NM) & vbNewLine  ''''''''''''''''''''''''
         End If
         ModSize = Len(ModString$)
         pSOProc = StartOfProcsPos
         If pSOProc < ModSize Then      ' To allow for no procs
            FirstIn = 0
            Do
               pEOL = InStr(pSOProc, ModString$, Chr$(10))
               If pEOL - pSOProc - 1 < 1 Then Exit Do
               ' Collect line
               A$ = Mid$(ModString$, pSOProc, pEOL - pSOProc - 1)
               pSONextL = pEOL + 1
               
               For k = 1 To PriE1 ' Public Sub -- Friend Property
                  If InStr(1, A$, ProcName$(k)) = 1 Then
                     
                     EndString$ = "End Sub"
                     If InStr(1, ProcName$(k), "Function ") > 0 Then
                        EndString$ = "End Function"
                     ElseIf InStr(1, ProcName$(k), "Property ") > 0 Then
                        EndString$ = "End Property"
                     End If
                     ' Keep Proc start line
                     'FirstIn = 0
                     p2 = InStr(1, A$, "(")
                     If p2 = 0 Then Stop ' ERROR
                     p = InStrRev(A$, " ", p2)
                     B$ = Mid$(A$, p + 1, p2 - p)
                     ' B$= ProcName(
                     ' NEED:::-- to SORT Controls Procs from others
                     ' Check if ProcName used
                     ' Publics/Friends?/Statics?  throughout project  ' ModCollection$
                     ' Privates in mod ' ModString$
                     ProcStartLine$ = Space$(10) & B$ & vbNewLine
                     pund = InStrRev(B$, "_")
                     If pund > 0 Then
                        C$ = Left$(B$, pund)   ' Form_ ,  cmdB_ etc OR  Name_ of Name_Name
                        E$ = Mid$(B$, pund + 1)
                        If C$ = "Form_" Or C$ = "MDIForm_" Or C$ = "UserControl_" Or C$ = "PropertyPage_" Then
                           C$ = Left$(B$, pund - 1) & " "   ' ie Form^
                        Else
                           C$ = C$  ' ie mnuList_   OTHER_[mmm]
                        End If
                     End If
                     If LISTTYPE = 0 Then ' CONTROL PROC NAMES
                        ' Ctrl procs  Check if Name in ModCtrlName$(NM, k)
                        If pund > 0 Then  ' _ possible Ctrl
                           ' Check if C$ in ModCtrlName$(NM, MaxNumCtrls)
                           For i = 1 To MaxNumCtrls
                              If ModCtrlName$(NM, i) = "" Then Exit For
                              If InStr(1, ModCtrlName$(NM, i), C$) Then
                                 ' Check E$ Events
                                 For j = 1 To NumEventTypes
                                    If E$ = EventType$(j) Then Exit For
                                 Next j
                                 If j < NumEventTypes + 1 Then
                                    NProcs = NProcs + 1
                                    RT$ = RT$ & ProcStartLine$  ''''''''''''''''''''''''
                                    Exit For
                                 End If
                              End If
                           Next i
                        Else  ' non ctrl
                           ' IGNORE
                        End If
                     Else  ' LISTTYPE = 1, 2 or 3  ' 1 Non-Ctrl procs,  2 Callers, 3 Unused Non-Ctrl procs
                        i = MaxNumCtrls + 1      ' Force lister for pund = 0 (ie no _)
                        If pund <> 0 Then
                           ' _ Possible ctrl- not wanted
                           ' Check if C$ in ModCtrlName$(NM, MaxNumCtrls)
                           For i = 1 To MaxNumCtrls
                              If ModCtrlName$(NM, i) = "" Then ' End ofList
                                 i = MaxNumCtrls + 1
                                 Exit For
                              End If
                              If InStr(1, ModCtrlName$(NM, i), C$) Then
                                 ' Check E$ Events
                                 For j = 1 To NumEventTypes
                                    If E$ = EventType$(j) Then Exit For
                                 Next j
                                 If j < NumEventTypes + 1 Then  ' Ctrl event found ignore
                                    Exit For
                                 End If
                              End If
                           Next i
                        End If
                           
                        If i = MaxNumCtrls + 1 Then   ' No Ctrl Event found
                           '=========================================================
                           NProcs = NProcs + 1
                           Select Case LISTTYPE
                           Case 1   ' NON-CONTROL PROC NAMES
                              RT$ = RT$ & ProcStartLine$  ''''''''''''''''''''''''
                           Case 2   ' NON-CONTROL PROC CALLERS
                              C$ = Trim$(ProcStartLine$)
                              C$ = Left$(C$, Len(C$) - 3)      ' clip (crlf
                              If FirstIn = 0 Then RT$ = RT$ & ModName$(NM) & vbNewLine  '''''''''''''''
                              FirstIn = FirstIn + 1
                              
                              RT$ = RT$ & Space$(2) & "## " & A$ & " ##" & vbNewLine '''''''''''''''
                              RT$ = RT$ & Space$(8) & "CALLED FROM:-" & vbNewLine   '''''''''''''''
                              If Left$(A$, 3) = "Pri" Or Left$(A$, 3) = "Fri" Then ' Private
                                 LOOK_IN_MODSTRING A$, C$, ModName$(NM), "Pri", LISTTYPE
                              Else  ' Public
                                 TempModString$ = ModString$   ' Since LOOK_IN_MODCOLLECTION gets new ModString$s
                                 TempStartOfProcsPos = StartOfProcsPos
                                 LOOK_IN_MODCOLLECTION A$, C$, LISTTYPE
                                 ModString$ = TempModString$
                                 StartOfProcsPos = TempStartOfProcsPos
                                 TempModString$ = ""
                              End If
                           Case 3   ' UNUSED NON-CONTROL PROCS
                              C$ = Trim$(ProcStartLine$)
                              C$ = Left$(C$, Len(C$) - 3)      ' clip (crlf
                              If Left$(A$, 3) = "Pri" Or Left$(A$, 3) = "Fri" Then ' Private
                                 LOOK_IN_MODSTRING A$, C$, ModName$(NM), "Pri", LISTTYPE
                                 If NUM_PROCS_USED_IN_MOD = 0 Then
                                    TOTALUNUSED = TOTALUNUSED + 1
                                    RT$ = RT$ & Space$(2) & "## " & A$ & " ##" & vbNewLine '''''''''''''''
                                    'RT$ = RT$ & Space$(8) & "UNUSED !!" & vbNewLine
                                 End If
                              Else  ' Public
                                 TempModString$ = ModString$   ' Since LOOK_IN_MODCOLLECTION gets new ModString$s
                                 TempStartOfProcsPos = StartOfProcsPos
                                 LOOK_IN_MODCOLLECTION A$, C$, LISTTYPE
                                 If ALL_NUM_PROCS_USED = 0 Then
                                    TOTALUNUSED = TOTALUNUSED + 1
                                    RT$ = RT$ & Space$(2) & "## " & A$ & " ##" & vbNewLine '''''''''''''''
                                    'RT$ = RT$ & Space$(8) & "UNUSED !!" & vbNewLine
                                 End If
                                 ModString$ = TempModString$
                                 StartOfProcsPos = TempStartOfProcsPos
                                 TempModString$ = ""
                              End If
                           End Select
                           '=========================================================
                        End If
                     End If
                     ' Get End of Proc
                     pEOProc = InStr(pSOProc, ModString$, Chr$(10) & EndString$)
                     ' pEOProc ->LF End Sub/Function/Property
                     pEOProc = pEOProc + Len(EndString$) + 2      ' -> next LF
                     Exit For
                  End If
               
               Next k
               If k = PriE1 + 1 Then Exit Do  ' No more ProcNames found -> next module
               pSOProc = pEOProc + 1
               If pSOProc >= ModSize Then Exit Do      ' ModSize = Len(ModString$)
            Loop   ' Loop thru all lines of ModString$ seeking ProcName$(k)
         End If  ' If pSOProc < ModSize Then
      End If  ' If LenB(ModString$) > 0
   Next NM  ' Next Module
   ' Return with RT$
End Sub

Public Sub LOOK_IN_MODSTRING(ByVal PName$, ByVal SName$, MName$, Typ$, LISTTYPE As Long)
' Find Private XREFS
' IN:
' PName$=Full Proc line ' for eliminating calls to itself
' SName$=Proc name searched for in ModString$
' MName$ ModName$()

' ModString$
' RT$ -> RTMod text

Dim A$
Dim p As Long
Dim p2 As Long
Dim p3 As Long
Dim i As Long
Dim j As Long
Dim k As Long
' ModString$ positions
Dim pSOProc As Long     ' Ptr Start of Proc
Dim pEOL As Long        ' Ptr Start of Line
Dim pSONextL As Long    ' Ptr Start of Next Line
Dim pEOProc As Long     ' Ptr End of Proc
Dim j4 As Long
Dim ModSize As Long

Dim NM As Long
Dim TotNumDims As Long
Dim NumDimsInProc As Long
Dim NProcs As Long
Dim EndString$
Dim USED As Long
Dim TOTALUNUSED As Long
Dim ProcString$
Dim ProcStartLine$
Dim FirstIn As Long
Dim Bef$
Dim Aft$
Dim Last$
Dim skip As Long
   ModSize = Len(ModString$)
   pSOProc = StartOfProcsPos
   A$ = Mid$(ModString$, StartOfProcsPos, 30)
   If pSOProc < ModSize Then      ' To allow for no procs
      NUM_PROCS_USED_IN_MOD = 0
      FirstIn = 0
      Do
         pEOL = InStr(pSOProc, ModString$, Chr$(10))
         If pEOL - pSOProc - 1 < 1 Then Exit Do
         ' Collect line
         A$ = Mid$(ModString$, pSOProc, pEOL - pSOProc - 1)
         pSONextL = pEOL + 1
         For k = 1 To PriE1 ' Public Sub -- Friend Property
            If InStr(1, A$, ProcName$(k)) = 1 Then
               
               EndString$ = "End Sub"
               If InStr(1, ProcName$(k), "Function ") > 0 Then
                  EndString$ = "End Function"
               ElseIf InStr(1, ProcName$(k), "Property ") > 0 Then
                  EndString$ = "End Property"
               End If
               pEOProc = InStr(pSOProc, ModString$, Chr$(10) & EndString$)
               pEOProc = pEOProc + Len(EndString$) + 2      ' -> next LF
               ProcString$ = Chr$(10) & Mid$(ModString$, pSOProc, pEOProc - pSOProc)
               p = 0
               skip = 0
               Do
                  p = InStr(p + 1, ProcString$, SName$)
                  If p > 0 Then
                     If A$ <> PName$ Or skip = 1 Then
                        If p = 1 Then
                           Bef$ = Chr$(10)
                        Else
                           Bef$ = Mid$(ProcString$, p - 1, 1)
                        End If
                        Aft$ = Mid$(ProcString$, p + Len(SName$), 1)
                        Last$ = ""
                        If TEST_CHAR(Bef$, Aft$, Last$) Then
                           NUM_PROCS_USED_IN_MOD = NUM_PROCS_USED_IN_MOD + 1
                           ALL_NUM_PROCS_USED = ALL_NUM_PROCS_USED + 1
                           If LISTTYPE <> 3 Then RT$ = RT$ & Space$(8) & MName$ & " :: " & A$ & vbNewLine  '''''''''''''''
                           Exit Do  ' Only need to know if called at least once
                        End If
                     Else
                        skip = 1 ' maybe recursive
                     End If
                  Else     ' None or no more SName$ to test in ProcString$
                     Exit Do
                  End If
                  If p + 1 >= Len(ProcString$) Then Exit Do
               Loop
               Exit For
            End If
         Next k
         If k = PriE1 + 1 Then Exit Do  ' No more ProcNames found -> return
         pSOProc = pEOProc + 1
         If pSOProc >= ModSize Then Exit Do      ' ModSize = Len(ModString$)
      Loop   ' Loop thru all Procs in ModString$ seeking SName$
   End If
   If LISTTYPE = 2 And Typ$ = "Pri" And NUM_PROCS_USED_IN_MOD = 0 Then
      RT$ = RT$ & Space$(8) & "UNUSED (or Class/Dsr/REv IT)" & vbNewLine
   End If
   ' Return RT$
End Sub

Public Sub LOOK_IN_MODCOLLECTION(ByVal PName$, ByVal SName$, LISTTYPE As Long)
' Find Public XREFS

' IN:
' PName$=Full Proc line ' for eliminating calls to itself
' SName$=Proc name searched for in ModString$

Dim NM As Long
ALL_NUM_PROCS_USED = 0
   For NM = 1 To NumMods
      ' Extract module string from collection
      ' into ModString$ & calc StartOfProcsPos
      EXTRACT_MODSTRING ModName$(NM)
      If LenB(ModString$) > 0 Then
         LOOK_IN_MODSTRING PName$, SName$, ModName$(NM), "Pub", LISTTYPE
      End If
   Next NM
   If LISTTYPE = 2 And ALL_NUM_PROCS_USED = 0 Then
      RT$ = RT$ & Space$(8) & "UNUSED (or Class/Dsr/REv IT)" & vbNewLine
   End If
End Sub

Public Sub LIST_DRCs_MODCOLLECTION(LISTTYPE As Long, NProcs As Long, TotNumDims As Long, TOTALUNUSED As Long)
' LISTTYPE = 0 ' List Procs Dims ReDims & Consts DRCs
' LISTTYPE = 1 ' List Procs & Unused Dims

' Called from:
' mnuListProcs_Details_Click
' mnuUNUB_Click

Dim A$, B$, C$
Dim p As Long
Dim p2 As Long
Dim p3 As Long
Dim k As Long

' ModString$ positions
Dim pSOProc As Long  ' Ptr Start of Procs
Dim pEOL As Long     ' Ptr End of Line
Dim pSOL As Long     ' Ptr Start of Line
Dim pSONextL As Long ' Ptr Start of Next Line
Dim ModSize As Long  ' = Len(ModString$)

Dim NM As Long
Dim NumDRCsInProc As Long
Dim EndString$
Dim UNUSED As Long
Dim ProcString$
Dim ProcStartLine$
Dim FirstIn As Long
Dim Flag As Long

ReDim DRCStore$(50)
ReDim DRCType(50)

'p2 = UBound(DRCStore$(), 1)
'p3 = UBound(DRCType(), 1)

   ReDim NumProcs(ProcSub3)
   RT$ = ""
   TotNumDims = 0
   TOTALUNUSED = 0
   NProcs = 0
   For NM = 1 To NumMods
      ' Extract module string from collection
      ' into ModString$ & calc StartOfProcsPos
      EXTRACT_MODSTRING ModName$(NM)
      
      ' READ line at a time from ModString$ into A$
      ' & display declarations in ListProcs
      If LenB(ModString$) > 0 Then     ' Else goto next module
         Form1.Label3 = " " & ModName$(NM) & " "
         Form1.Label3.Refresh
         
         '###################################################################
         RT$ = RT$ & ModName$(NM) & vbNewLine  ''''''''''''''''''''''''
         ModSize = Len(ModString$)
         pSOProc = StartOfProcsPos
         If pSOProc < ModSize Then      ' To allow for no procs
            Do
               pEOL = InStr(pSOProc, ModString$, Chr$(10))
               If pEOL - pSOProc - 1 < 1 Then Exit Do
               ' Collect line
               A$ = Mid$(ModString$, pSOProc, pEOL - pSOProc - 1)
               pSOL = pEOL + 1
               NumDRCsInProc = 0
               For k = 1 To PriE1 ' Public Sub -- Friend Property
                  If InStr(1, A$, ProcName$(k)) = 1 Then
                     ' PROC FOUND
                     ' PROCESS PROC
                     ' Keep Proc start line
                     FirstIn = 0
                     ProcStartLine$ = Space$(10) & A$ & vbNewLine
                     EndString$ = "End Sub"
                     If InStr(1, ProcName$(k), "Function ") > 0 Then
                        EndString$ = "End Function"
                     ElseIf InStr(1, ProcName$(k), "Property ") > 0 Then
                        EndString$ = "End Property"
                     End If
                     ' READ ALL LINES IN PROC
                     ' Get & store Dim/ReDim/Const Vars
                     Do
                        pSONextL = InStr(pSOL, ModString$, Chr$(10)) + 1
                        If pSONextL <= pSOL Then Exit Do
                        B$ = Mid$(ModString$, pSOL, pSONextL - pSOL)  ' Includes CRLF
                        ' End Sub/Function/Property
                        If InStr(1, B$, EndString$) = 1 Then Exit Do
                        Flag = -1
                        If InStr(1, B$, "Dim ") = 1 Then
                           Flag = 0
                        ElseIf InStr(1, B$, "ReDim ") = 1 Then
                           Flag = 1
                        ElseIf InStr(1, B$, "Const ") = 1 Then
                           Flag = 2
                        End If
                        If Flag = 0 Or (LISTTYPE = 0 And (Flag = 1 Or Flag = 2)) Then
                           ' For LISTTYPE = 0 also Collect Redims & Consts
                           ' Extract & Store Dim/ReDim/Consts variables for the proc
                           NumDRCsInProc = NumDRCsInProc + 1
                           TotNumDims = TotNumDims + 1
                           If NumDRCsInProc > UBound(DRCStore$, 1) Then
                              ReDim Preserve DRCStore$(UBound(DRCStore$, 1) + 20)
                              ReDim Preserve DRCType(UBound(DRCType, 1) + 20)
                           End If
                           EXTRACT_DRC_NAMES B$, C$
                           DRCStore$(NumDRCsInProc) = C$
                           DRCType(NumDRCsInProc) = Flag ' 0 Dims, 1 ReDims, 2 Consts
                        End If
                        pSOL = pSONextL
                        If pSOL >= ModSize Then Exit Do         ' ModSize = Len(ModString$)
                     Loop   ' Loop thru whole Proc to find any Dim/ReDim/Const
                     ' pSOL -> EndString$
                     pSOL = InStr(pSOL, ModString$, Chr$(10))
                     ' pSOL -> End of Proc (Could be End of ModString$)
                     If NumDRCsInProc > 0 Then
                        ' pSOProc = position of Start of Proc
                        If LISTTYPE = 0 Then  ' List Procs Dims ReDims & Consts
                           NProcs = NProcs + 1
                           FirstIn = FirstIn + 1
                           If FirstIn = 1 Then RT$ = RT$ & ProcStartLine$ ''''''''''''''
                           ' Print DRCs
                           For p = 1 To NumDRCsInProc
                              A$ = "Dim "
                              If DRCType(p) = 1 Then A$ = "ReDim "
                              If DRCType(p) = 2 Then A$ = "Const "
                              RT$ = RT$ & Space$(12) & A$ & DRCStore$(p) & vbNewLine ''''''''''''''
                           Next p
                           ' & number of DRCs in proc
                           RT$ = RT$ & Str$(NumDRCsInProc) & " DRCs" & vbNewLine '''''''''''''''
                        Else  ' LISTTYPE = 1  List Procs & Unused Dims
                           ' Search for use of Dims 0,1, > 1
                           ProcString$ = Mid$(ModString$, pSOProc, pSOL - pSOProc + 1)
                           UNUSED = 0
                           'FIND_UNUSED_DIMS Chr$(10) & ProcString$, DRCStore$(), NumDRCsInProc, UNUSED
                           FIND_UNUSED_DIMS Chr$(10) & ProcString$, NumDRCsInProc, UNUSED
                           If UNUSED > 0 Then
                              NProcs = NProcs + 1
                              FirstIn = FirstIn + 1
                              If FirstIn = 1 Then RT$ = RT$ & ProcStartLine$ '''''''''''''
                              TOTALUNUSED = TOTALUNUSED + UNUSED
                              For p = 1 To UNUSED
                                 RT$ = RT$ & Space$(12) & "Dim " & DRCStore$(p) & vbNewLine ''''''''''''
                              Next p
                           End If   ' If UNUSED > 0
                        End If   ' If LISTTYPE = 0 Then
                     End If   ' If NumDRCsInProc > 0
                     Exit For
                  
                  End If   ' If InStr(1, A$, ProcName$(k)) = 1
               Next k   ' Public Sub -- Friend Property
               If k = PriE1 + 1 Then Exit Do
               ' pSOL -> End of Proc (Could be End of ModString$)
               pSOProc = pSOL + 1
               If pSOProc >= ModSize Then Exit Do      ' ModSize = Len(ModString$)
            Loop   ' Loop thru whole ModString$ with next ProcName$(k)
         End If  ' If pSOProc < ModSize Then
      '###################################################################
      End If  ' If LenB(ModString$) > 0
   Next NM  ' Next Module
   ' Return with RT$
   Erase DRCStore$()
   Erase DRCType()
End Sub

Public Sub GET_VARS(NM As Long, IPrint As Integer, KeyWord$, NumPubPriVars As Long, NEnumTypeStarts As Long, _
   Filt1A As Long, Filt2A As Long, Filt1B As Long, Filt2B As Long, aFilter As Boolean, _
   ProcSubscript As Long, NumInMod As Long)
   
Dim A$, B$
Dim p As Long
Dim p1 As Long
Dim p2 As Long
Dim p3 As Long
Dim p4 As Long
Dim k As Long
Dim pSOL As Long  ' Ptr Start of Line
Dim pEOL As Long  ' Ptr End of Line
'Dim NM As Long
'Dim NumInMod As Long
'Dim Filt1A As Long, Filt2A As Long
'Dim Filt1B As Long, Filt2B As Long
'Dim aFilter As Boolean

'Dim ProcSubscript As Long
   
      ' READ line at a time from ModString$ into A$
      ' & display declarations in ListProcs
      ' Find prelims
      If LenB(ModString$) > 0 Then
         pSOL = 1    ' Also Start of Code initially
         Do
            pEOL = InStr(pSOL, ModString$, Chr$(10))
            If pEOL - pSOL + 1 < 1 Then Exit Do
            A$ = Mid$(ModString$, pSOL, pEOL - pSOL + 1)
            B$ = A$
            p = InStr(1, A$, KeyWord$)
            If p > 0 Then
               If aFilter Then
                  ' Skip unwanted Public or Privates
                  For k = 1 To ProcSub1   ' Public Sub -- Public,Private
                     If ProcName$(k) <> "" Then ' Breaks
                        Select Case k
                        ' Filter out Publics & Privates other than
                        ' "Public " or "Private "
                        Case Filt1A To Filt2A
                           p1 = InStr(1, A$, ProcName$(k))
                           If p1 = 1 Then Exit For
                        ' Filter out Publics & Privates other than
                        ' "Public " or "Private "
                        Case Filt1B To Filt2B
                           p1 = InStr(1, A$, ProcName$(k))
                           If p1 = 1 Then Exit For
                        End Select
                     End If
                  Next k
               End If
               ' Test if filtered out A$ detected
               If (aFilter And k = ProcSub1 + 1) Or Not aFilter Then ' Get name
                  If Left$(A$, Len(KeyWord$)) = KeyWord$ Then
                     p1 = Len(KeyWord$)
                     Select Case KeyWord$
                     Case "Private ", "Public ", "Public Event "
                        p2 = InStr(p1 + 1, A$, "(")
                        If p2 = 0 Then
                           p2 = InStr(p1 + 1, A$, " ")
                           If p2 = 0 Then
                              p2 = InStr(p1 + 1, A$, Chr$(13))
                           End If
                        End If
                     Case "Public Sub ", "Private Sub ", _
                          "Public Function ", "Private Function ", _
                          "Public Property ", "Private Property ", _
                          "Static Sub ", "Friend Sub ", _
                          "Static Function ", "Friend Function ", _
                          "Static Property ", "Friend Property "
                        
                        p2 = InStr(p1 + 1, A$, "(") ' Bracket after name
                        If p2 = 0 Then
                           p2 = InStr(p1 + 1, A$, " ")
                           If p2 = 0 Then
                              p2 = InStr(p1 + 1, A$, Chr$(13))
                           End If
                        Else
                           p2 = p2 - 1
                        End If
                     Case Else
                        p2 = InStr(p1 + 1, A$, " ")
                        If p2 = 0 Then
                           p2 = InStr(p1 + 1, A$, "(")
                           If p2 = 0 Then
                              p2 = InStr(p1 + 1, A$, Chr$(13))
                           End If
                        End If
                     End Select
                     ' ModName$(NM)
                     ' VarName = Mid$(A$, p1 + 1, p2 - p1 )
                     A$ = Mid$(A$, p1 + 1, p2 - p1)

                     ' Condition name
                     ' Remove any CR
                     If Right(A$, 1) = Chr$(13) Then
                        A$ = Left$(A$, Len(A$) - 1)
                     End If
                     ' Remove any trailng spaces
                     If Right$(A$, 1) = Chr$(32) Then
                        A$ = Trim$(A$)
                     End If
                     NumPubPriVars = NumPubPriVars + 1
                     NumInMod = NumInMod + 1
                     If NumPubPriVars > UBound(PubPrivStore$) Then
                        ReDim Preserve ModNameStore$(UBound(ModNameStore$) + 20)
                        ReDim Preserve PubPrivStore$(UBound(PubPrivStore$) + 20)
                     End If
                     ModNameStore$(NumPubPriVars) = ModName$(NM)
                     PubPrivStore$(NumPubPriVars) = A$
                     If IPrint <> 1 Then
                        B$ = ModName$(NM)
                        If Len(B$) > 28 Then
                           B$ = Left$(B$, 11) & String$(2, ".") & Right$(B$, 15)
                        End If
                        RT$ = RT$ & B$ & Space$(30 - Len(B$)) & A$ & vbNewLine '''''''''''''''
                     End If
                     If KeyWord$ = "Public Enum " Or KeyWord$ = "Private Enum " Then
                        NEnumTypeStarts = NEnumTypeStarts + 1
                        ' List Enum elements
                        Do
                           pSOL = pEOL + 1
                           pEOL = InStr(pSOL, ModString$, Chr$(10))
                           A$ = Mid$(ModString$, pSOL, pEOL - pSOL + 1)
                           B$ = A$
                           If InStr(1, A$, "End Enum") = 0 Then
                              p2 = InStr(1, A$, " ")
                              If p2 = 0 Then
                                 p2 = InStr(1, A$, Chr$(13))
                              End If
                              A$ = Left$(A$, p2 - 1)
                              NumPubPriVars = NumPubPriVars + 1
                              NumInMod = NumInMod + 1
                              If NumPubPriVars > UBound(PubPrivStore$) Then
                                 ReDim Preserve ModNameStore$(UBound(ModNameStore$) + 20)
                                 ReDim Preserve PubPrivStore$(UBound(PubPrivStore$) + 20)
                              End If
                              ModNameStore$(NumPubPriVars) = ModName$(NM)
                              PubPrivStore$(NumPubPriVars) = A$
                              If IPrint <> 1 Then
                                 RT$ = RT$ & Space$(32) & A$ & vbNewLine  ''''''''''''''''
                              End If
                              ' Flag Enum elements
                              PubPrivStore$(NumPubPriVars) = "*" & PubPrivStore$(NumPubPriVars)
                           Else
                              Exit Do
                           End If
                        Loop
                     End If
                     If KeyWord$ = "Public Type " Or KeyWord$ = "Private Type " Then
                        NEnumTypeStarts = NEnumTypeStarts + 1
                        ' List Type elements
                        Do
                           pSOL = pEOL + 1
                           pEOL = InStr(pSOL, ModString$, Chr$(10))
                           A$ = Mid$(ModString$, pSOL, pEOL - pSOL + 1)
                           B$ = A$
                           If InStr(1, A$, "End Type") = 0 Then
                              p2 = InStr(1, A$, " ")
                              If p2 = 0 Then
                                 p2 = InStr(1, A$, Chr$(13))
                              End If
                              A$ = "." & Left$(A$, p2 - 1)
                              'A$ = Left$(A$, p2 - 1)
                              NumPubPriVars = NumPubPriVars + 1
                              NumInMod = NumInMod + 1
                              If NumPubPriVars > UBound(PubPrivStore$) Then
                                 ReDim Preserve ModNameStore$(UBound(ModNameStore$) + 20)
                                 ReDim Preserve PubPrivStore$(UBound(PubPrivStore$) + 20)
                              End If
                              ModNameStore$(NumPubPriVars) = ModName$(NM)
                              PubPrivStore$(NumPubPriVars) = A$
                              If IPrint <> 1 Then
                                 RT$ = RT$ & Space$(32) & A$ & vbNewLine ''''''''''''''
                              End If
                           Else
                              Exit Do
                           End If
                        Loop
                     End If
                  End If
               End If
            End If   ' If p > 0 Then
            pSOL = pEOL + 1
            If ProcSubscript >= PubS2 Then
               If pSOL >= StartOfProcsPos Then Exit Do      ' To get next mod
            Else  ' ProcSubscript < PubS2
               If pSOL >= Len(ModString$) Then Exit Do      ' To get next mod
            End If
         Loop
      End If

End Sub

