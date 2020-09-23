Attribute VB_Name = "HiLitRTB"
'quick richtext rtf highlighting demo hack using api for psc
'bugbyter 02-2003

'ripped from vbaccelerator
'http://vbaccelerator.nuwebhost.com/codelib/richedit/richedit.htm
'watch their source code for PARAFORMAT2 (paragraph settings) and all the constants!
'...there is so much more!
'this is just a quick hack as i found its missing on PSC (and the whole web), don't have time for more...

'now found more on psc, too:
'http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=38434&lngWId=1

Option Explicit

Public Const LF_FACESIZE = 32
Public Const WM_USER = &H400
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Public Const CFM_BACKCOLOR = &H4000000
Public Const SCF_SELECTION = &H1

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

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
    szFaceName(0 To LF_FACESIZE - 1) As Byte    ' 58
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
