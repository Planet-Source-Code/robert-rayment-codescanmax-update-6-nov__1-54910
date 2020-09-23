Attribute VB_Name = "basColor"
' basColor.bas

' Based on code by Will Barden
'
Option Explicit

Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Const COL_KEYWORD = vbRed '&H800000    ' dark blue
Private Const COL_COMMENT = &H8000&     ' middle green
Private Const CHAR_COMMENT = "'"        ' comment line char

Private Type WORD_TYPE
    Text As String  ' word to be colored
    Color As Long   ' color to color the word :)
End Type

Private Type LETTER_TYPE
    Start As Long   ' first time the letter appears in the list
    Finish As Long  ' last time the letter appears in the list
End Type

Private Words() As WORD_TYPE
Private Letters() As LETTER_TYPE
Private Strings() As String
Private sText As String

Public Sub InitKeyWords()
'  Builds the arrays of keywords, then builds
'  an alphabetical index of the array to aid
'  searching later on.
Dim k As Long
   ' initialize the array of words
   ReDim Words(0 To 108)
   ' NB if Words added then change 108
    
   Words(0).Text = "Access"
   Words(1).Text = "AddressOf"
   Words(2).Text = "Alias"
   Words(3).Text = "And"
   Words(4).Text = "As"
   Words(5).Text = "Boolean"
   Words(6).Text = "Byte"
   Words(7).Text = "Currency"
   Words(8).Text = "Date"
   Words(9).Text = "Double"
   Words(10).Text = "Integer"
   Words(11).Text = "Long"
   Words(12).Text = "Object"
   Words(13).Text = "Single"
   Words(14).Text = "String"
   Words(15).Text = "Variant"
   Words(16).Text = "BF"
   Words(17).Text = "Base"
   Words(18).Text = "Begin"
   Words(19).Text = "Binary"
   Words(20).Text = "ByRef"
   Words(21).Text = "ByVal"
   Words(22).Text = "CBool"
   Words(23).Text = "CByte"
   Words(24).Text = "CCur"
   Words(25).Text = "CDate"
   Words(26).Text = "CDbl"
   Words(27).Text = "CInt"
   Words(28).Text = "CLng"
   Words(29).Text = "CSgn"
   Words(30).Text = "CStr"
   Words(31).Text = "CVar"
   Words(32).Text = "Call"
   Words(33).Text = "Case"
   Words(34).Text = "Circle"
   Words(35).Text = "Close"
   Words(36).Text = "Const"
   Words(37).Text = "Declare"
   Words(38).Text = "Dim"
   Words(39).Text = "Do"
   Words(40).Text = "Friend"
   Words(41).Text = "Each"
   Words(42).Text = "Else"
   Words(43).Text = "ElseIf"
   Words(44).Text = "Empty"
   Words(45).Text = "End"
   Words(46).Text = "Enum"
   Words(47).Text = "Error"
   Words(48).Text = "Event"
   Words(49).Text = "Exit"
   Words(50).Text = "Explicit"
   Words(51).Text = "False"
   Words(52).Text = "For"
   Words(53).Text = "Function"
   Words(54).Text = "Get"
   Words(55).Text = "GoTo"
   Words(56).Text = "If"
   Words(57).Text = "In"
   Words(58).Text = "Input"
   Words(59).Text = "Is"
   Words(60).Text = "LBound"
   Words(61).Text = "Let"
   Words(62).Text = "Lib"
   Words(63).Text = "Like"
   Words(64).Text = "Line"
   Words(65).Text = "Local"
   Words(66).Text = "Lock"
   Words(67).Text = "Loop"
   Words(68).Text = "Mod"
   Words(69).Text = "New"
   Words(70).Text = "Next"
   Words(71).Text = "Not"
   Words(72).Text = "Nothing"
   Words(73).Text = "On"
   Words(74).Text = "Open"
   Words(75).Text = "Option"
   Words(76).Text = "Optional"
   Words(77).Text = "Or"
   Words(78).Text = "Output"
   Words(79).Text = "Preserve"
   Words(80).Text = "Print"
   Words(81).Text = "Private"
   Words(82).Text = "Property"
   Words(83).Text = "Public"
   Words(84).Text = "RaiseEvent"
   Words(85).Text = "Random"
   Words(86).Text = "ReDim"
   Words(87).Text = "Read"
   Words(88).Text = "Resume"
   Words(89).Text = "Seek"
   Words(90).Text = "Select"
   Words(91).Text = "Set"
   Words(92).Text = "Step"
   Words(93).Text = "Sub"
   Words(94).Text = "Then"
   Words(95).Text = "To"
   Words(96).Text = "True"
   Words(97).Text = "Type"
   Words(98).Text = "TypeOf"
   Words(99).Text = "UBound"
   Words(100).Text = "Until"
   Words(101).Text = "Wend"
   Words(102).Text = "While"
   Words(103).Text = "With"
   Words(104).Text = "Write"
   
   Words(105).Text = "Static"
   Words(106).Text = "ParamArray"
   Words(107).Text = "Any"
   Words(108).Text = "Erase"
   
   For k = 0 To UBound(Words(), 1)
          Words(k).Color = COL_KEYWORD
   Next k
    
   ' IE Pre-Do ComboSort & BuildIndex
   CombSort Words
   ' build the index of letter positions
   BuildIndex
End Sub

Public Sub ColorRTB(frm As Form, RTB As RichTextBox)
' Enter with RTB filled
' Public LineCount As Long
Dim lStart As Long
Dim lFinish As Long
Dim Text As String

Dim xp As Single
Dim xd As Single
   If LineCount > 0 Then
   
      ' split the text into lines and color them one by one
      LockWindowUpdate RTB.hwnd
      RTB.Visible = False
      basColor.sText = RTB.Text
      lStart = 1
      
      aColoringDone = False
      
      frm.picPB.Visible = True
      frm.picPB.DrawWidth = 3
      frm.picPB.Cls
      xd = frm.picPB.Width / LineCount
      
      Do While lStart <> 2 And lStart < Len(sText)
         ' find the end of this line
         lFinish = InStr(lStart + 1, sText, vbCrLf)
         If lFinish = 0 Then lFinish = Len(sText)
         ' color it
         DoColor RTB, lStart, lFinish
         ' move up to get the next line
         lStart = lFinish + 2
         DoEvents
         If aColoringDone Then Exit Do
         
         frm.picPB.PSet (xp, 1) '-(xp, 1)
         xp = xp + xd
      
      Loop
   End If
   frm.picPB.Visible = False
   
   ' reset the cursor
   RTB.SelStart = 0
   RTB.Visible = True
   LockWindowUpdate 0&
End Sub

Public Sub DoColor(RTB As RichTextBox, ByVal lStart As Long, ByVal lFinish As Long)
'  This routine colors a single line of text within the RTB. It will
'  split each line up into words using the custom split function (SplitWords),
'  then match each word against the list of keywords.

' NB RR Not 100% but FAST !!

Dim sWords()    As String
Dim sLine       As String
Dim sChar       As String
Dim lCurPos     As Long
Dim lIndex      As Long
Dim lColor      As Long
Dim lPos        As Long
Dim lPos2       As Long
Dim lCom        As Long
Dim i           As Long

   ' grab the line
   sLine = Trim$(Mid$(sText, lStart, lFinish - lStart))
   ' remove the EOL
   sLine = RemoveEOL(sLine)
   ' remove the quotes so they're not colored
   sLine = RemoveStrings(sLine)
   ' split the line into words using our custom function
   sWords = SplitWords(sLine)
   
   ' check each word against the list
   lCurPos = 1
   ' search for each word in the array
   For i = LBound(sWords) To UBound(sWords)
      
      If Trim$(sWords(i)) <> "" Then
   
          ' check for comment in the middle of a line
          If Left$(sWords(i), 1) = CHAR_COMMENT Then
          
            ' color the rest of the line
            RTB.SelStart = InStr(lStart, sText, sWords(i)) - 1
            RTB.SelLength = Len(sWords(i))
            RTB.SelColor = COL_COMMENT
          
          Else
      
             ' its a normal keyword - so color it
             ' first get the array positions from
             ' the index
             ' Get first char of sWords(i)
             sChar = Left$(LCase$(sWords(i)), 1)
             ' if we've got a valid alphabetic char
             If sChar <> "" Then
                ' convert this char to an index in the letters array
                lIndex = Asc(sChar) - 97
                ' if the index is a valid one - this
                ' means that the text is a word, so
                ' we should try to color it
                If lIndex >= 0 And lIndex < UBound(Letters) Then
                  ' color the word, passing the index parameters
                  lColor = GetColor(sWords(i), _
                              Letters(lIndex).Start, _
                              Letters(lIndex).Finish)
                  ' if a color was returned - color the word
                  If lColor Then
                     ' locate the word in the line
                     lPos = InStr(lStart + lCurPos - 1, sText, sWords(i)) - 1
                     'lPos = InStr(lStart + lCurPos, sText, sWords(i)) - 1
                     If lPos >= 0 Then
                        RTB.SelStart = lPos 'InStr(lStart + lCurPos - 1, sText, sWords(i)) - 1
                        RTB.SelLength = Len(sWords(i))
                        RTB.SelColor = lColor
                     End If
                  End If
                End If
                'lCurPos = lCurPos + Len(sWords(i))
            Else
               'lCurPos = lCurPos + Len(sWords(i)) + 1
            End If ' sChar <> ""
         End If ' CHAR_COMMENT
         'lCurPos = lCurPos + Len(sWords(i)) + 1
      Else
          'lCurPos = lCurPos + 1
      End If ' sWords(i) <> ""
      
      ' move the current position within the line on
      lCurPos = lCurPos + Len(sWords(i)) + 1
       
   Next i
End Sub

Private Function GetColor(ByVal sWord As String, _
                          ByVal Lo As Long, _
                          ByVal Hi As Long) As Long
'  Searches the Words array for a match using a standard
'  binary search algorithm, using the Lo and Hi params
'  as starting points.
Dim lHi As Long
Dim lLo As Long
Dim lMid As Long
    
Dim k As Long
    
    ' standard binary search the words array
    ' return the color if a match is found
    ' ??????????
   GetColor = 0
   lLo = Lo
   lHi = Hi
   Do While lHi >= lLo
      lMid = (lLo + lHi) \ 2
      If LCase$(Words(lMid).Text) = LCase$(sWord) Then
          GetColor = Words(lMid).Color
          Exit Do
      End If
      'If LCase$(Words(lMid).Text) > LCase$(sWord) Then  ' Error I think in original
      If (Words(lMid).Text) > (sWord) Then
          lHi = lMid - 1
      Else
          lLo = lMid + 1
      End If
   Loop
'    ' Alternative to binary search
'    For k = Lo To Hi
'        If LCase$(Words(k).Text) = LCase$(sWord) Then
'            GetColor = Words(k).Color
'            Exit For
'        End If
'    Next k
End Function

Private Function SplitWords(ByVal sText As String) As String()
'  Since splitting a line into words by a single
'  character is not acceptable because we have to
'  take several end of word characters into account,
'  this routine was written.
'  It searches through the string from left to right
'  and locates the nearest word break char from a list
'  then splits at that word.
Dim i As Long, lPos As Long
Dim sWords() As String
Dim sWordBreaks(0 To 8) As String
Dim lBreakPoints() As Long
Dim lBreak As Long
    
'    ' list of word break characters
   sWordBreaks(0) = " "
   sWordBreaks(1) = "("
   sWordBreaks(2) = ")"
   sWordBreaks(3) = "<"
   sWordBreaks(4) = ">"
   sWordBreaks(5) = "."
   sWordBreaks(6) = ","
   sWordBreaks(7) = "="
   sWordBreaks(8) = CHAR_COMMENT ' comments
   ReDim lBreakPoints(UBound(sWordBreaks))
   
   ' get them words!
   ReDim sWords(0)
   lPos = 1
   Do
      ' locate the word break points
      For i = 0 To UBound(sWordBreaks)
         lBreakPoints(i) = InStr(lPos, sText, sWordBreaks(i))
      Next i
      
      ' now work out which is closest
      lBreak = Len(sText) + 1
      For i = 0 To UBound(lBreakPoints)
        If lBreakPoints(i) <> 0 Then
            If lBreakPoints(i) < lBreak Then lBreak = lBreakPoints(i)
        End If
      Next i
   
      ' now split out the word
      ' if no break point was found, then we've
      ' hit the end of the line, so add all the rest
      If lBreak = Len(sText) + 1 Then
         sWords(UBound(sWords)) = Mid$(sText, lPos)
      Else
        ' add this word - first check for a comment
        If Mid$(sText, lBreak, 1) = CHAR_COMMENT Then
           ' first add the word
           sWords(UBound(sWords)) = Mid$(sText, lPos, lBreak - lPos)
           ' then add the rest as a comment
           ReDim Preserve sWords(UBound(sWords) + 1)
           sWords(UBound(sWords)) = Mid$(sText, lBreak)
           ' now return and exit
           SplitWords = sWords
           Exit Function
        Else
            sWords(UBound(sWords)) = Mid$(sText, lPos, lBreak - lPos)
        End If
      End If
      ReDim Preserve sWords(UBound(sWords) + 1)
   
      ' move the pointer on a bit
      lPos = lBreak + 1
      
      ' setup the exit condition
      If lPos >= Len(sText) Then Exit Do
   
   Loop

   ' return the array
   SplitWords = sWords
End Function

Private Function RemoveEOL(ByVal sText As String) As String
'  Removes leading and trailing vbCrLf from strings
Dim sTmp As String
    ' remove leading or trailing vbCrLf from the string
    sTmp = sText
    If Left$(sTmp, 2) = vbCrLf Then
        sTmp = Right$(sTmp, Len(sTmp) - 2)
    End If
    If Right$(sTmp, 2) = vbCrLf Then
        sTmp = Left$(sTmp, Len(sTmp) - 2)
    End If
    RemoveEOL = sTmp
End Function

Private Function RemoveStrings(ByVal sText As String) As String
'  Removes any quoted strings from the text, but only
'  those that aren't within comments of course.
Dim lCom As Long
Dim lPos As Long
Dim lPos2 As Long

   lCom = InStr(1, sText, CHAR_COMMENT)
   lPos = InStr(1, sText, Chr$(34))
      If lPos < lCom Or lCom = 0 Then
         Do While lPos <> 0
            ' find the end " char to make a pair
            lPos2 = InStr(lPos + 1, sText, Chr$(34))
            If lPos2 <> 0 Then
               ' we've found a pair, so remove it
               sText = Mid$(sText, 1, lPos - 1) & Mid$(sText, lPos2 + 1)
               ' find the next starting " avoiding
               ' comments within strings
               lCom = InStr(lPos2 + 1, sText, CHAR_COMMENT)
               lPos = InStr(lPos2 + 1, sText, Chr$(34))
               If lPos > lCom Then Exit Do
            Else
                Exit Do
            End If
         Loop
      End If
   ' return
   RemoveStrings = sText
End Function

Private Sub BuildIndex()
'  Takes the Words array and constructs an alphabetical
'  index which it puts into the Letters array.
'  Each item in the letters array accounts for a letter
'  in the alphabet - Letters(0) = "a".
'  The .Start property is the Index in the Words array
'  at which that letter starts, and the finish is the
'  same. The purpose of this is to get Hi and Lo params
'  for the GetColor (a standard binary search algorithm).
'  This saves several loops round the algorithm.
Dim i As Long, j As Long, k As Long
Dim sChar As String
Dim bStart As Boolean

   ' go through each letter in the alphabet
   ReDim Letters(25)
   For i = 0 To 25
      ' get the current char
      sChar = Chr$(i + 97)
      ' find the first and last instances of the letter
      For j = LBound(Words) To UBound(Words)
         If Left$(LCase$(Words(j).Text), 1) = sChar Then
            If Not bStart Then
                ' found the start
                bStart = True
                Letters(i).Start = j
            End If
            ' if we've hit the end of the list
            If j = UBound(Words) Then
                Letters(i).Finish = j
                Exit Sub
            End If
         Else
            ' its a different char
            If bStart Then
                ' we've found the end
                Letters(i).Finish = j - 1
                bStart = False
                Exit For
            End If
            ' see if we've gone too far -
            ' there are no words beginning with
            ' this letter in the list
            If Left$(LCase$(Words(j).Text), 1) > sChar Then
                Exit For
            End If
         End If
      Next j
   Next i

End Sub


'//--[CombSort]------------------------------------------------------------//
'  Will's comments
'  This is a standard comb sort - you could replace
'  this with any other sorting algorithm, I just prefer
'  this one because a) i wrote it :), and b) it performs
'  well across all ranges of input arrays - it makes
'  no assumptions about how sorted the array already
'  is, because it doesn't matter.
'  The comb sort is a slight variation on the bubblesort,
'  and i know what you're thinking - ewwww, bubble sorts -
'  but you'd be wrong, the comb is only fractionally
'  slower than a quicksort... so enjoy!
'  for more on the combsort, read here:
'  http://yagni.com/combsort/index.php
'  http://cs.clackamas.cc.or.us/molatore/cs260Spr01/combsort.htm
'
Private Sub CombSort(Arr() As WORD_TYPE)
Dim i As Long, j As Long, t As WORD_TYPE
Dim swapped As Boolean
Dim gap As Long
   
   gap = UBound(Arr)
   
   Do
      gap = (gap * 10) \ 13
      If gap = 9 Or gap = 10 Then gap = 11
      If gap < 1 Then gap = 1
      
      swapped = False
      For i = 0 To UBound(Arr) - gap
         j = i + gap
         If Arr(i).Text > Arr(j).Text Then
            LSet t = Arr(j)
            LSet Arr(j) = Arr(i)
            LSet Arr(i) = t
            swapped = True
         End If
      Next i
      
      If (gap = 1) And (Not swapped) Then Exit Do
   Loop

   ' Optionally Write Words Code Lines
   '   eg Words(0).Text = "Private"
'    Open "ComboSort.txt" For Output As #2
'    For i = LBound(Arr) To UBound(Arr)
'      Print #2, "Words(" & Trim$(Str$(i)) & ").Text = " & Chr$(34) & Arr(i).Text & Chr$(34)
'    Next i
'    Close #2
End Sub

