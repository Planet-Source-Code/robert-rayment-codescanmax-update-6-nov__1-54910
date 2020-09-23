Attribute VB_Name = "Stripper"

' Stripper.bas by Robert Rayment

' STRIP down any module - removing all unnecessary characters,
' simplifying multiple line statements and adding Public/Private
' as necessary.


' IN:
' ModSpec$ = ModFileSpec$
' StripHeader=0 leave in module header
' StripHeader=1 STRIP out all before code

' Clean out:-
' comments if not in quotes,
' blank lines,
' leading spaces,
' multiple spaces replaced by 1 space if not in quotes.
' Remove continuation chars (ie space_CRLF).

' Replace :space by CRLF (ie when : folllowed by a space but not followed
'  by CRLF (ie Label:) or other character (Header) and not in quotes.)

' Replace Global by Public (even if in quotes!!).
' Optionally remove header stuff.
' IF header removed then
'  place vars with commas on separate lines (ie Public/Private/Dim var1,var2, etc)
'  & replace Dim with Private in declaration section

' OUT:
'Public ModString$ = stripped module ( with or without header)
'Public StartOfCodePos As Long
'Public StartOfProcsPos As Long

Option Explicit
Option Base 1



Private bArray() As Byte
Private bArray2() As Byte
Private FSize As Long
Private BYT As Byte
'Private bprevB As Byte
Private NQuote As Long
'Private NComment As Long
Private NPos As Long
'Private NCR As Long
Private NLF As Long
'Private NContin As Long
Private NSpaces As Long
'Private NColon As Long
'Private kpos As Long

Private k As Long:         Private p As Long
Private A$, B$, C$
Private j1 As Long
Private j2 As Long
Private pcomma As Long
Private pub As Long
Private pri As Long
Private pdim As Long
Private pb1 As Long
Private pb2 As Long
'Private repstring$
Private aSkip As Boolean
Private aValid As Boolean
Private aRefindStartOfProcsPos As Boolean
Private Ext$
Private lena As Long

' TEST VARS
Private bb As Byte
Private pq As Long
Private pq2 As Long
Private p2  As Long
Private LenModSring As Long

Public Sub STRIP(ModNum As Long, ModSpec$, StripHeader As Integer)
' IN:
' ModSpec$ = ModFileSpec$
' StripHeader=0 leave in module header
' StripHeader=1 STRIP out all before code

' Clean out:-
' comments if not in quotes,
' blank lines,
' leading spaces,
' multiple spaces replaced by 1 space if not in quotes.
' Remove continuation chars (ie space_CRLF).
' Replace :space by CRLF (ie when : not followed by CRLF (ie Label:) and not in quotes),
' Replace Global by Public (even if in quotes!!).
' Optionally remove header stuff.
' IF header removed then
'  place vars with commas on separate lines (ie Public/Private/Dim var1,var2, etc)
'  & replace Dim with Private in declaration section

' OUT:
'Public ModString$ = stripped module ( with or without header)
'Public StartOfCodePos As Long
'Public StartOfProcsPos As Long

' TEST --------
LLL:
'--------------

   On Error GoTo MODREADERR
   
   If LenB(ModSpec$) < 1 Then Exit Sub
   
   Open VBPDir$ & ModSpec$ For Binary Access Read As #1
   FSize = LOF(1)
   If FSize = 0 Then
      Close
      MsgBox "Can't read " & VBPDir$ & ModFileSpec$, vbCritical, " Reading VB Files"
      Exit Sub
   End If
   
   ReDim bArray(FSize)
   Get #1, , bArray()
   Close
   ReDim bArray2(FSize)
   NLF = 0
'========================================================================
   ' Clears:
   ' leading spaces
   ' blank lines
   ' comments
   ' replace (space_CRLF) by a space
   ' Input in bArray(FSize)
   
   ' Uses Mod scoped:-  p, BYT, A$, lena, FSize, j1, j2, NPos
   
   InitialClearOut   ' Output in bArray2(), Npos bytes
   
   ReDim Preserve bArray2(NPos)
   ReDim bArray(NPos)
   FSize = NPos
   
   NPos = 0
   NQuote = 0
   NSpaces = 0
'========================================================================
   ' Replace multiple spaces with 1 space unless in quotes
   
   ' Uses Mod scoped:-  BYT, FSize, NQuote, NPos, NSpaces
   
   ReplaceMultipleSpaces   ' Output in bArray(), Npos bytes

   ReDim Preserve bArray(NPos)
   
'========================================================================
   ' Replace :space by CRLF if not in quotes  ' Will not alter NPos
   
   ' Uses  Mod scoped:- NPos, NQuote
   
   SeparateMultiLines
   
   ReDim Preserve bArray(NPos)   ' Output in bArray(), Npos bytes
   
'========================================================================
'========================================================================
   ' COPY bArray(Npos) to ModString$
   ModString$ = Space$(NPos)
   CopyMemory ByVal ModString$, bArray(1), NPos
   Erase bArray()
   Erase bArray2()
   
'===============================================================
   ' Replace Global by Public, even if in quotes !!
   ModString$ = Replace(ModString$, "Global ", "Public ", 1)
   
'===============================================================
   ' Find start of code - StartOfCodePos
   
   ' Uses  Mod scoped:- p, A$
   '       Public StartOfCodePos, ModString$
   
   FindStartOfCode   ' Out: ModString$
   
   FindCtrlNames ModNum, ModSpec$ ' From ModString$ before stripping -> ModCtrlName$(NumMods,#)
   
'===============================================================
   ' Optionally STRIP header lines
   If StripHeader = 1 Then
      ModString$ = Mid$(ModString$, StartOfCodePos)
      StartOfCodePos = 1
      
      If LenB(ModString$) = 0 Then
         If LCase$(Right$(ModSpec$, 3)) = "frm" Then
            ModString$ = "Private Sub Form_Load()" & vbCrLf & "End Sub" & vbCrLf
         Else
            ModString$ = "Private Sub Dummy()" & vbCrLf & "End Sub" & vbCrLf
         End If
      End If
      '===============================================================
      ' For stripped code only:-
      ' replace  Public/Private/Dim var1 [As Type], var 2 [As Type], etc
      '      by  Public/Private/Dim var1 [As Type]
      '          Public/Private/Dim var2 [As Type]
      ' etc
   
      ' Uses  Mod scoped:- j1, j2, A$, k, aSkip
      ' Public ModString$
      
      SeparateCommaDefines
   End If
   
'===============================================================
   ' Find start of Procs - StartOfProcsPos
   ' Also needed if StripHeader = 1
   
   ' Uses  Mod scoped:- p, A$
   '       Public StartOfCodePos, ModString$
   
   FindStartOfProcs
   
'===============================================================
   If StripHeader = 1 Then ' ie header removed
   
      ' Replace Dim by Private in declaration section  'nb "Dim " -> "Private "?
      ' Replace Declare Sub  or Declare Function by
      '  Public Declare Sub  or Public Declare Function in declaration section
      
      ' Uses  Mod scoped:- A$, NPos
      ' Public StartOfCodePos, ModString$
      
      DeclarationReplacer
      
      ' Replace Sub or Function at start of a line by
      '  Public Sub or Public Function
      
      ' Uses  Mod scoped:- A$, NPos, LenModSring
      ' Public StartOfCodePos, ModString$
      ProcedureReplacer
      
      '===============================================================
      ' Replace any Enum, Type, Const, Property in FRM or BAS, CLS, CTL files, by
      '  Private/Public Enum, Public Type, Private/Public Const, Public Property
      
      ' Uses  Mod scoped:- p, aRefindStartOfProcsPos
      
      aRefindStartOfProcsPos = False
      Ext$ = UCase$(Right$(ModSpec$, 3))
      Select Case Ext$
      Case "FRM", "DSR", "PAG"
         FRM_Replacer            ' Returns aRefindStartOfProcsPos True/False
      Case "BAS", "CLS", "CTL"
         BASCLSCTL_Replacer      ' Returns aRefindStartOfProcsPos True/False
      End Select
      
      '===============================================================
      If aRefindStartOfProcsPos Then
         ' Find NEW start of Procs -  - StartOfProcsPos
         
         ' Uses  Mod scoped:- p, A$
         '       Public StartOfCodePos, ModString$
         
         FindStartOfProcs
      End If
      
   End If
'===============================================================
'   If (StartOfProcsPos = 1 And InStr(1, ModString$, "Friend")) Or _
'      (StartOfProcsPos > 1 And InStr(1, ModString$, Chr$(10) & "Friend")) Then
'      Dim Form As Form
'      MsgBox "Sorry Friend not dealt with", vbCritical, "Stripper.bas"
'      ModCollection$ = ""
'      ModString$ = ""
'      DoEvents
'      ' Make sure all forms cleared
'      For Each Form In Forms
'         Unload Form
'         Set Form = Nothing
'      Next Form
'      End
'   End If
'===============================================================
StartOfProcsPos = StartOfProcsPos
Exit Sub
'========
MODREADERR:
   MsgBox "Error in Stripper.bas: STRIP", vbCritical, "CodeScan"
   Close
   ModString$ = ""
End Sub

Private Sub InitialClearOut()
Dim pqcnt As Long
Dim i As Long
' Clears:
' leading spaces
' blank lines
' comments
' replace (space_CRLF) by a space
   NPos = 0
   j1 = 1
   j2 = j1
   NPos = 1
   Do
      Do ' Find LF @ EOL
         If bArray(j2) = 10 Then Exit Do
         j2 = j2 + 1
      Loop
      ' Extract line from bArray()
      lena = j2 - j1 + 1
      A$ = Space$(lena) ' Includes CRLF
      CopyMemory ByVal A$, bArray(j1), lena
      
      A$ = LTrim$(A$)     ' Clear leading spaces
      
      If A$ <> Chr$(13) & Chr$(10) Then   ' ie skip blank line
         p = InStr(1, A$, " _" & Chr$(13))
         If p > 0 Then
'Debug.Print A$
            If Mid$(A$, p - 1, 1) = "(" Then
               A$ = Left$(A$, p - 1)            ' ie replace (sp_ by (
            Else
               A$ = Left$(A$, p)                ' ie replace sp_ by a space
               If Mid$(A$, p - 1, 1) = " " Then
                  A$ = Left$(A$, p - 1) & Mid$(A$, p + 1)
               End If
            End If
         End If
         
         If Left$(A$, 1) <> "'" Then         ' ie skip comment at BOL
         If Left$(A$, 4) <> "Rem " Then      ' ie skip Rem @ BOL
         '-------------------
            ' Clear comments
            p = InStr(1, A$, "'")
            If p > 0 Then
               If InStr(1, A$, Chr$(34)) = 0 Then   ' Comment only no quotes
                  A$ = Left$(A$, p - 1)
                  A$ = Trim(A$) & Chr$(13) & Chr$(10)
   
               Else
                  ' Comment(s) & quote(s)
                  ' Start counting quotes
                  pqcnt = 0
                  For i = 1 To Len(A$)
                     C$ = Mid$(A$, i, 1)
                     If C$ = Chr$(34) Then
                        pqcnt = pqcnt + 1
                     ElseIf C$ = Chr$(39) Then
                        If pqcnt = 0 Then
                           A$ = Trim$(Mid$(A$, 1, i - 1)) & Chr$(13) & Chr$(10)
                           Exit For
                        End If
                     End If
                     If pqcnt = 2 Then pqcnt = 0
                  Next i
                  A$ = Trim$(A$)
               End If
            End If
            
            CopyMemory bArray2(NPos), ByVal A$, Len(A$)
            NPos = NPos + Len(A$)
         
         '-------------------
         End If
         End If
      
      End If
      
      j1 = j1 + lena ' NB original length of A$ in bArray()
      j2 = j1
      If j1 > FSize Then Exit Do
   Loop
   NPos = NPos - 1
End Sub

Private Sub ReplaceMultipleSpaces()
' Replace multiple spaces with 1 space unless in quotes
   For k = 1 To FSize
      BYT = bArray2(k)
      If BYT = 34 Then        '  "
         NQuote = NQuote + 1
         NPos = NPos + 1
         bArray(NPos) = bArray2(k)
         If NQuote = 2 Then NQuote = 0 ' "" "" will leave  space""
      ElseIf BYT = 32 Then
         If NQuote = 0 Then
            NSpaces = NSpaces + 1
            If NSpaces = 1 Then
               NPos = NPos + 1
               bArray(NPos) = bArray2(k)
            Else
               NSpaces = NSpaces - 1
            End If
         Else
            NSpaces = 0
            NPos = NPos + 1
            bArray(NPos) = bArray2(k)
         End If
      Else
         NSpaces = 0
         NPos = NPos + 1
         bArray(NPos) = bArray2(k)
      End If
   Next k
End Sub

Private Sub SeparateMultiLines()
' Replace :space by CRLF if not in quotes  ' Will not alter NPos
   NQuote = 0
   For k = 1 To NPos:
      If bArray(k) = 34 Then  '  "
         NQuote = NQuote + 1
         If NQuote = 2 Then NQuote = 0
      End If
      If NQuote = 0 Then
         If bArray(k) = 58 Then     ' :^
            If k < NPos Then
               If bArray(k + 1) <> 13 Then   ' ie not Label:
                  If bArray(k + 1) = 32 Then   'skip : no space (in header)
                     bArray(k) = 13: bArray(k + 1) = 10
                  End If
               End If
            End If
         End If
      End If

   Next k
End Sub

Private Sub FindCtrlNames(ModN As Long, ModSpec$)
'IN: ModN the Module number as in ListMods order
'    ModSpec$ the Module filename
Dim Ext$
Dim A$, B$
Dim p As Long, p2 As Long
Dim N As Long
Dim pSOL As Long    ' Ptr Start of Line
Dim pEOL As Long    ' Ptr End of Line
Dim BeginCount As Long

N = 0
' From ModString$ before any stripping -> ModCtrlName$(NumMods,#)
' Have ModSpec$  .frm, .ctl etc
   p = InStrRev(ModSpec$, ".")
   Ext$ = Mid$(ModSpec$, p + 1)
   Select Case LCase$(Ext$)
   Case "frm", "ctl", "pag"
      pSOL = 1
      Do
         pEOL = InStr(pSOL, ModString$, Chr$(10))
         A$ = Mid$(ModString$, pSOL, pEOL - pSOL + 1)
         If InStr(1, A$, "Begin ") = 1 Then
            BeginCount = BeginCount + 1
            ' Now find ctrl name
            p = InStr(1, A$, ".")
            If p = 0 Then
               End    ' ERROR
            End If
            p2 = InStr(p, A$, " ")
            
            B$ = Mid$(A$, p + 1, p2 - p - 1)
            'If B$ = "Form" Or B$ = "MDIForm"  Or B$ = "UserControl" Or B$ = "PropertyPage" Then
               B$ = Mid$(A$, p + 1)    ' CtrlType ^ Name
            'Else
            '   B$ = Mid$(A$, p2 + 1)  ' CtrlName
            'End If
            
            p = InStr(1, B$, Chr$(13))
            If p > 0 Then
               B$ = Left$(B$, p - 1)
            End If
            B$ = Trim$(B$)
            B$ = B$ & "_"     '??
            N = N + 1
            
            If N = 1 Then
               If N > MaxNumCtrls Then
                  MaxNumCtrls = N
                  ReDim Preserve ModCtrlName$(NumMods, MaxNumCtrls)
               End If
               ModCtrlName$(ModN, N) = B$
            
            Else
               If B$ <> ModCtrlName$(ModN, N - 1) Then   ' Avoid Ctrl Indexes > 0
                  If N > MaxNumCtrls Then
                     MaxNumCtrls = N
                     ReDim Preserve ModCtrlName$(NumMods, MaxNumCtrls)
                  End If
                  ModCtrlName$(ModN, N) = B$
               Else
                  N = N - 1
               End If
            End If
         ElseIf InStr(1, A$, "End" & Chr$(13)) = 1 Then
            BeginCount = BeginCount - 1
            If BeginCount = 0 Then Exit Do
         End If
         
         If InStr(1, A$, "Attribute VB_") Then Exit Do ' Long stop
         
         pSOL = pEOL + 1
         If pSOL >= Len(ModString$) Then Exit Do
      Loop
   Case Else   ' bas, cls, dsr
      N = N + 1
      If N > MaxNumCtrls Then
         MaxNumCtrls = N
         ReDim Preserve ModCtrlName$(NumMods, MaxNumCtrls)
      End If
      ModCtrlName$(ModN, N) = ""
   End Select
   
End Sub

Private Sub FindStartOfCode()
' Find start of code - usually beginning of declarations section
   StartOfCodePos = 1
   p = InStr(1, ModString$, "Attribute VB_")
   StartOfCodePos = p
   Do
      p = InStr(StartOfCodePos + 1, ModString$, "Attribute VB_")
      If p = 0 Then Exit Do
      A$ = Mid$(ModString$, p - 1, 1)
      If A$ = Chr$(10) Then   ' LF
         StartOfCodePos = p
      Else  ' string "Attribute VB_" elsewhere in code
         Exit Do
      End If
   Loop
   p = InStr(StartOfCodePos + 1, ModString$, Chr$(13))
   StartOfCodePos = p + 2
End Sub

Private Sub FindStartOfProcs()
' NEW MENU DONE

' Find start of procs ie first line after Declarations section
   NPos = Len(ModString$)
   StartOfProcsPos = NPos
   'If StartOfProcsPos = 0 Then Exit Sub
   For k = 1 To PriE1  ' Public Sub -- Friend Property
      p = InStr(1, ModString$, ProcName$(k))
      If p > 0 Then
         If p > 1 Then A$ = Mid$(ModString$, p - 1, 1)
         j1 = InStr(p, ModString$, Chr$(10))
         B$ = Mid$(ModString$, p, j1 - p)
         If p = 1 Or A$ = Chr$(10) Then
            If p < StartOfProcsPos Then
               StartOfProcsPos = p
            End If
         End If
         Do Until p = 0
            If p + 1 > NPos Then Exit Do
            p = InStr(p + 1, ModString$, ProcName$(k))
            If p > 0 Then
               A$ = Mid$(ModString$, p - 1, 1)
               ' test that ProcName$(k) is at BOL
               If A$ = Chr$(10) Then
                  If p < StartOfProcsPos Then
                     StartOfProcsPos = p
                  End If
               End If
            End If
         Loop
      End If
   Next k
   
   For k = ProcSub1 + 1 To ProcSub2 ' Sub , Function , Property
      p = InStr(1, ModString$, ProcName$(k))
      If p > 0 Then
         If p > 1 Then A$ = Mid$(ModString$, p - 1, 1)
         j1 = InStr(p, ModString$, Chr$(10))
         B$ = Mid$(ModString$, p, j1 - p)
         If p = 1 Or A$ = Chr$(10) Then
            If p < StartOfProcsPos Then
               StartOfProcsPos = p
            End If
         End If
         Do Until p = 0
            If p + 1 > NPos Then Exit Do
            p = InStr(p + 1, ModString$, ProcName$(k))
            If p > 0 Then
               A$ = Mid$(ModString$, p - 1, 1)
               ' test that ProcName$(k) is at BOL
               If A$ = Chr$(10) Then
                  If p < StartOfProcsPos Then
                     StartOfProcsPos = p
                  End If
               End If
            End If
         Loop
      End If
   Next k

End Sub

Private Sub FRM_Replacer()
Dim A$
' If required ADD Private before Const, Type or Public before Enum, Property
   p = InStr(1, ModString$, "Const ")
   If p = 1 Then
      ModString$ = "Public " & ModString$
      aRefindStartOfProcsPos = True
   Else
      p = InStr(1, ModString$, Chr$(10) & "Const ")
      If p > 0 Then
      If p < StartOfProcsPos Then   ' Since in Procs Const xx = allowed but not with Public or Private
         A$ = Left$(ModString$, StartOfProcsPos - 1)
         A$ = Replace(A$, Chr$(10) & "Const", Chr$(10) & "Private Const")
         ModString$ = A$ & Mid$(ModString$, StartOfProcsPos)
         A$ = ""
         aRefindStartOfProcsPos = True
      End If
      End If
   End If
   
'   'Illegal
'   p = InStr(1, ModString$, "Type ")
'   If p = 1 Then
'      ModString$ = Replace(ModString$, "Type", "Private Type")
'      aRefindStartOfProcsPos = True
'   Else
'      p = InStr(1, ModString$, Chr$(10) & "Type ")
'      If p > 0 Then
'         ModString$ = Replace(ModString$, Chr$(10) & "Type", Chr$(10) & "Private Type")
'         aRefindStartOfProcsPos = True
'   End If
   
   p = InStr(1, ModString$, "Enum ")
   If p = 1 Then
      ModString$ = "Public " & ModString$
      aRefindStartOfProcsPos = True
   Else
      p = InStr(1, ModString$, Chr$(10) & "Enum ")
      If p > 0 Then
         ModString$ = Replace(ModString$, Chr$(10) & "Enum", Chr$(10) & "Public Enum")
         aRefindStartOfProcsPos = True
      End If
   End If
   
   p = InStr(1, ModString$, "Property ")
   If p = 1 Then
      ModString$ = "Public " & ModString$
      aRefindStartOfProcsPos = True
   Else
      p = InStr(1, ModString$, Chr$(10) & "Property ")
      If p > 0 Then
         ModString$ = Replace(ModString$, Chr$(10) & "Property", Chr$(10) & "Public Property")
         aRefindStartOfProcsPos = True
      End If
   End If
End Sub

Private Sub BASCLSCTL_Replacer()
Dim A$
' If required ADD Public before Enum, Type, Const, Property
   p = InStr(1, ModString$, "Enum ")
   If p = 1 Then
      ModString$ = "Public " & ModString$
      aRefindStartOfProcsPos = True
   Else
      p = InStr(1, ModString$, Chr$(10) & "Enum ")
      If p > 0 Then
         ModString$ = Replace(ModString$, Chr$(10) & "Enum", Chr$(10) & "Public Enum")
         aRefindStartOfProcsPos = True
      End If
   End If
   
   p = InStr(1, ModString$, "Type ")
   If p = 1 Then
      A$ = Mid$(ModString$, p, 10)
      ModString$ = "Public " & ModString$
      aRefindStartOfProcsPos = True
   Else
      p = InStr(1, ModString$, Chr$(10) & "Type ")
      If p > 0 Then
         A$ = Mid$(ModString$, p, 10)
         ModString$ = Replace(ModString$, Chr$(10) & "Type", Chr$(10) & "Public Type")
         aRefindStartOfProcsPos = True
      End If
   End If
   
   p = InStr(1, ModString$, "Const ")
   If p = 1 Then
      ModString$ = "Public " & ModString$
   Else
      p = InStr(1, ModString$, Chr$(10) & "Const ")
      If p > 0 Then
         ModString$ = Replace(ModString$, Chr$(10) & "Const", Chr$(10) & "Public Const")
         aRefindStartOfProcsPos = True
      End If
   End If
   
   p = InStr(1, ModString$, "Property ")
   If p = 1 Then
      ModString$ = "Public " & ModString$
   Else
      p = InStr(1, ModString$, Chr$(10) & "Property ")
      If p > 0 Then
         ModString$ = Replace(ModString$, Chr$(10) & "Property", Chr$(10) & "Public Property")
         aRefindStartOfProcsPos = True
      End If
   End If

   p = InStr(1, ModString$, "Declare ")
   If p = 1 Then
      ModString$ = "Public " & ModString$
   Else
      p = InStr(1, ModString$, Chr$(10) & "Declare ")
      If p > 0 Then
         ModString$ = Replace(ModString$, Chr$(10) & "Declare", Chr$(10) & "Public Declare")
         aRefindStartOfProcsPos = True
      End If
   End If
End Sub

Private Sub SeparateCommaDefines()
' Replace  any Public/Private/Dim var1 [As Type], var2(1 to 2, 1 to 4) [As Type], var 3 [As Type], etc
' by separate lines
   j1 = 1
   Do
      j2 = InStr(j1, ModString$, Chr$(10))
      If j2 - j1 + 1 < 1 Then
         A$ = Mid$(ModString$, j1)
         Exit Do
      End If
      A$ = Mid$(ModString$, j1, j2 - j1 + 1)
      p = Asc(A$)
      
      ' Skip past other than Public/Private/Dim in
      ' declarations section
      aSkip = False
      For k = 1 To PriE2    ' Public Sub - Private WithEvents
         If ProcName$(k) <> "" Then
            p = InStr(1, A$, ProcName$(k))
            If p > 0 Then Exit For   ' Done
         End If
      Next k
      If k < PriE2 + 1 Then
         aSkip = True
      End If
      
      If Not aSkip Then
         ' Leaves any Public, Private, Dim, (Global has been replaced by Public)
         pub = InStr(1, A$, "Public ")
         pri = InStr(1, A$, "Private ")
         pdim = InStr(1, A$, "Dim ")
         If pub = 1 Or pri = 1 Or pdim = 1 Then
            ' Any ( in this line
            pb1 = InStr(1, A$, "(")
            If pb1 > 0 Then
               pb2 = InStr(pb1, A$, ")")
               B$ = Left$(A$, pb1) & Mid$(A$, pb2)
               If j1 > 1 Then
                  ModString$ = Left$(ModString$, j1 - 1) & B$ & Mid$(ModString$, j1 + Len(A$))
                  C$ = Mid$(ModString$, j1)
                  A$ = B$
               Else
                  ModString$ = B$ & Mid$(ModString$, j1 + Len(A$))
                  A$ = B$
               End If
            End If
            ' Any commas in this line
            pcomma = InStr(1, A$, ", ")
            If pcomma > 0 Then   ' P/P/D vars1 [As Type], vars2 [As Type]
                                 ' Also Public VarName(1 To 20, 1 To 10), Var2(1 To 4)
               If pub > 0 Then
                  B$ = Replace(A$, ", ", vbNewLine & "Public ", 1)
               ElseIf pri > 0 Then
                  B$ = Replace(A$, ", ", vbNewLine & "Private ", 1)
               Else  ' pdim > 0
                  B$ = Replace(A$, ", ", vbNewLine & "Dim ", 1)
               End If
               ' Insert B$
               If j1 > 1 Then
                  ModString$ = Left$(ModString$, j1 - 1) & B$ & Mid$(ModString$, j1 + Len(A$))
               Else
                  ModString$ = B$ & Mid$(ModString$, j1 + Len(A$))
               End If
                  C$ = Mid$(ModString$, j1)
               j2 = InStr(j1, ModString$, Chr$(10))
            End If
         End If
      End If
      C$ = Mid$(ModString$, j1)  ' TEST
      j1 = j2 + 1
      If j1 >= Len(ModString$) - 1 Then Exit Do ' Done
   Loop
End Sub

Private Sub DeclarationReplacer()
      'If Len(ModString$) = 0 Then Exit Sub
      ' Replace Dim by Private in declaration section  'nb "Dim " -> "Private "?
      A$ = Left$(ModString$, StartOfProcsPos - 1)
      A$ = Replace(A$, "Dim ", "Private ", 1)
      ModString$ = A$ & Mid$(ModString$, StartOfProcsPos)
      A$ = ""
      If Len(ModString$) > NPos Then  ' Dims replaced Privates therefore
         ' find NEW start of Procs  +4 for each replacement
         StartOfProcsPos = StartOfProcsPos + (Len(ModString$) - NPos)
         NPos = Len(ModString$)
      End If
      
      ' Replace Declare Sub  or Declare Function by
      ' Public Declare Sub  or Public Declare Function in declaration section
      If Left$(ModString$, 12) = "Declare " Then
         ModString$ = "Public Declare " & ModString$
      End If
      If Len(ModString$) > NPos Then
         StartOfProcsPos = StartOfProcsPos + (Len(ModString$) - NPos)
         NPos = Len(ModString$)
      End If
      A$ = Left$(ModString$, StartOfProcsPos - 1)
      A$ = Replace(A$, Chr$(10) & "Declare ", Chr$(10) & "Public Declare ", 1)
      ModString$ = A$ & Mid$(ModString$, StartOfProcsPos)
      If Len(ModString$) > NPos Then
         StartOfProcsPos = StartOfProcsPos + (Len(ModString$) - NPos)
         NPos = Len(ModString$)
      End If
End Sub

Private Sub ProcedureReplacer()
Dim p As Long
      ' Replace Sub or Function at start of a line by
      ' Public Sub or Pubic Function
      If Left$(ModString$, 4) = "Sub " Then
         ModString$ = "Public " & ModString$
      End If
      If Left$(ModString$, 9) = "Function " Then
         ModString$ = "Public " & ModString$
      End If
      ' StartOfProcsPos unaltered
      LenModSring = Len(ModString$)
      If StartOfProcsPos = 1 Then
         ModString$ = Chr$(10) & ModString$
         A$ = ModString$
         A$ = Replace(A$, Chr$(10) & "Sub ", Chr$(10) & "Public Sub ", 1)
         A$ = Replace(A$, Chr$(10) & "Function ", Chr$(10) & "Public Function ", 1)
         ModString$ = Right$(A$, Len(A$) - 1)
         'ModString$ = Left$(ModString$, StartOfProcsPos - 1) & Right$(A$, Len(A$) - 1)
         'A$ = ModString$
         'A$ = Mid$(ModString$, StartOfProcsPos - 1)
         'A$ = Replace(A$, Chr$(10) & "Function ", Chr$(10) & "Public Function ", 1)
         'ModString$ = Left$(ModString$, StartOfProcsPos - 1) & Right$(A$, Len(A$) - 1)
         'ModString$ = Right$(ModString$, Len(ModString$) - 1)
      ElseIf StartOfProcsPos > 1 Then
         A$ = Mid$(ModString$, StartOfProcsPos - 1)
         A$ = Replace(A$, Chr$(10) & "Sub ", Chr$(10) & "Public Sub ", 1)
         ModString$ = Left$(ModString$, StartOfProcsPos - 1) & Right$(A$, Len(A$) - 1)
         A$ = Mid$(ModString$, StartOfProcsPos - 1)
         p = InStr(StartOfProcsPos - 1, ModString$, Chr$(10) & "Function ")
         A$ = Replace(A$, Chr$(10) & "Function ", Chr$(10) & "Public Function ", 1)
         ModString$ = Left$(ModString$, StartOfProcsPos - 1) & Right$(A$, Len(A$) - 1)
      End If
End Sub
