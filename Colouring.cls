VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Colouring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum BlockType
  btWord = 1
  btFromTo = 2
  btFromToEOL = 3
End Enum

Private KeyWordMain() As Variant
Private KeyWordColour() As Long
Private KeyWordMain2() As Variant
Private KeyWordColour2() As Long
Private KeyWordMain3() As Variant
Private KeyWordColour3() As Long

Private IsInitialized As Boolean

Public Sub ColourCode(ByRef rtb As Object, SelectRTB As Boolean)
  Dim cpos As Long

  If Not IsInitialized Then Exit Sub

  cpos = rtb.SelStart
  rtb.Visible = False 'Hiding the rtb eliminates flicker and problems with scrolling
  
  'Make all the text black
  rtb.SelStart = 0
  rtb.SelLength = Len(rtb.Text)
  rtb.SelColor = vbBlack

' Zaznaczanie wyrazów
  Colour_Word rtb

' Zaznaczanie od znaku do znaku
  Colour_CharToChar rtb

' Zaznaczanie od znaku do koñca linii
  Colour_CharToEOL rtb

  rtb.SelStart = cpos
  rtb.Visible = True
  If SelectRTB Then
    On Error Resume Next
    If rtb.Visible = True Then 'Could be hidden by being in a hidden frame.
      rtb.SetFocus
    End If
  End If
End Sub

Private Sub Colour_Word(ByRef rtb As Object)
  Dim myword As String
  Dim pos As Long
  Dim i As Long, j As Long
  Dim i_lo As Long, i_hi As Long, j_lo As Long, j_hi As Long
  Dim colour As Long
'  Dim t1 As Single, t2 As Single

'  t1 = Timer
  j_lo = LBound(KeyWordMain): j_hi = UBound(KeyWordMain)
  For j = j_lo To j_hi
    colour = KeyWordColour(j)
    'Loop through each word
    i_lo = LBound(KeyWordMain(j)): i_hi = UBound(KeyWordMain(j))
    For i = i_lo To i_hi
      myword = KeyWordMain(j)(i)
      pos = 0
      'Loop Through each occurence of the word
      Do Until rtb.Find(myword, pos, , rtfWholeWord) = -1
        rtb.SelColor = colour
        pos = rtb.SelStart + rtb.SelLength
      Loop
    Next i
  Next j
'  t2 = Timer
'  rtb.SelStart = 0
'  rtb.SelLength = 0
'  rtb.SelText = "Word:" & Format(t2 - t1, "##0.000") & vbCrLf
End Sub

Private Sub Colour_CharToChar(ByRef rtb As Object)
  Dim myword As String, myword2 As String
  Dim pos As Long, pos2 As Long
  Dim i As Long, j As Long
  Dim i_lo As Long, i_hi As Long, j_lo As Long, j_hi As Long
  Dim colour As Long
'  Dim t1 As Single, t2 As Single

'  t1 = Timer
  j_lo = LBound(KeyWordMain2): j_hi = UBound(KeyWordMain2)
  For j = j_lo To j_hi
    colour = KeyWordColour2(j)
    'Loop through each word
    i_lo = LBound(KeyWordMain2(j)): i_hi = UBound(KeyWordMain2(j))
    For i = i_lo To i_hi
      myword = Left(KeyWordMain2(j)(i), Len(KeyWordMain2(j)(i)) \ 2)
      myword2 = Right(KeyWordMain2(j)(i), Len(KeyWordMain2(j)(i)) \ 2)
      'Loop Through each occurence of the word
      pos = rtb.Find(myword, 0, , rtfNoHighlight)
      Do Until pos = -1
        rtb.SelStart = pos
        pos2 = rtb.Find(myword2, pos + 1, , rtfNoHighlight)
        rtb.SelLength = IIf(pos2 = -1, Len(rtb.Text), pos2) - pos + Len(myword2)
        rtb.SelColor = colour
        pos = rtb.Find(myword, pos + rtb.SelLength, , rtfNoHighlight)
      Loop
    Next i
  Next j
'  t2 = Timer
'  rtb.SelStart = 0
'  rtb.SelLength = 0
'  rtb.SelText = "CharToChar:" & Format(t2 - t1, "##0.000") & vbCrLf
End Sub

Private Sub Colour_CharToEOL(ByRef rtb As Object)
  Dim myword As String
  Dim pos As Long, pos2 As Long
  Dim i As Long, j As Long
  Dim i_lo As Long, i_hi As Long, j_lo As Long, j_hi As Long
  Dim colour As Long
'  Dim t1 As Single, t2 As Single

'  t1 = Timer
  j_lo = LBound(KeyWordMain3): j_hi = UBound(KeyWordMain3)
  For j = j_lo To j_hi
    colour = KeyWordColour3(j)
    'Loop through each word
    i_lo = LBound(KeyWordMain3(j)): i_hi = UBound(KeyWordMain3(j))
    For i = i_lo To i_hi
      myword = KeyWordMain3(j)(i)
      'Loop Through each occurence of the word
      pos = rtb.Find(myword, 0, , rtfNoHighlight)
      Do Until pos = -1
        rtb.SelStart = pos
        pos2 = rtb.Find(vbCrLf, pos + 1, , rtfNoHighlight)
        rtb.SelLength = IIf(pos2 = -1, Len(rtb.Text), pos2) - pos + 1
        rtb.SelColor = colour
        pos = rtb.Find(myword, pos + rtb.SelLength, , rtfNoHighlight)
      Loop
    Next i
  Next j
'  t2 = Timer
'  rtb.SelStart = 0
'  rtb.SelLength = 0
'  rtb.SelText = "CharToEOL:" & Format(t2 - t1, "##0.000") & vbCrLf
End Sub

Public Property Get Delimiter() As String
  Delimiter = " "
End Property

Public Sub Initialize(What As BlockType, WordList As Variant, WordColorList As Variant, Delimiter As String)
  Dim i As Long

  Select Case What
  Case btWord
    ReDim KeyWordMain(LBound(WordList) To UBound(WordList))
    ReDim KeyWordColour(LBound(WordList) To UBound(WordList))
    For i = LBound(WordList) To UBound(WordList)
      KeyWordMain(i) = Split(WordList(i), Delimiter)
      KeyWordColour(i) = WordColorList(i)
    Next i
  Case btFromTo
    ReDim KeyWordMain2(LBound(WordList) To UBound(WordList))
    ReDim KeyWordColour2(LBound(WordList) To UBound(WordList))
    For i = LBound(WordList) To UBound(WordList)
      KeyWordMain2(i) = Split(WordList(i), Delimiter)
      KeyWordColour2(i) = WordColorList(i)
    Next i
  Case btFromToEOL
    ReDim KeyWordMain3(LBound(WordList) To UBound(WordList))
    ReDim KeyWordColour3(LBound(WordList) To UBound(WordList))
    For i = LBound(WordList) To UBound(WordList)
      KeyWordMain3(i) = Split(WordList(i), Delimiter)
      KeyWordColour3(i) = WordColorList(i)
    Next i
  End Select
End Sub

Public Property Get SQL_Text() As String
  SQL_Text = "''"
End Property

Public Property Get SQL_CommentLine() As String
  SQL_CommentLine = "--"
End Property

Public Property Get SQL_CommentBlock() As String
  SQL_CommentBlock = "/**/"
End Property


Public Property Get SQL_DataTypes() As String
  SQL_DataTypes = _
    "integer char numeric datetime int varchar real bit binary nchar nvarchar"
End Property

Public Property Get SQL_KeyWords() As String
  SQL_KeyWords = _
    "SELECT FROM WHERE DROP CREATE ON TRIGGER WITH ENCRYPTION FOR INSERT DELETE UPDATE AS BEGIN END INTO VALUES EXEC EXECUTE UNION IS ORDER HAVING BY WHEN THEN ELSE INNER GROUP FULL COMPUTE APPEND SHAPE TO RELATE IF TABLE COLLATE ALTER NOCHECK ADD CONSTRAINT PRIMARY KEY CLUSTERED DEFAULT UNIQUE NONCLUSTERED INDEX FOREIGN REFERENCES CASCADE PROCEDURE DECLARE CURSOR FAST_FORWARD OPEN FETCH NEXT CLOSE DEALLOCATE SET OFF QUOTED_IDENTIFIER ANSI_NULLS WHILE FUNCTION RETURNS RETURN VIEW TOP PRECENT USE MAX MIN DATABASE NAME FILENAME SIZE FILEGROWTH SCROLL INSTEAD OF RAISERROR DESC ASC PERCENT IDENTITY_INSERT DISABLE"
End Property

Public Property Get SQL_Operators() As String
  SQL_Operators = _
    "AND OR NOT NULL BETWEEN OUTER CROSS JOIN EXISTS IN ALL"
End Property
Public Property Get SQL_Functions() As String
  SQL_Functions = _
    "CASE SUM LEFT CAST CONVERT ISNULL DATEDIFF DATEADD DAY YEAR MONTH GETDATE COUNT RIGHT OBJECTPROPERTY OBJECT_ID RTRIM LTRIM LOG REPLACE ROUND DATEPART SUBSTRING NULLIF CHARINDEX AVG SPACE LEN UPPER " & _
    "@@FETCH_STATUS, @@IDENTITY"
End Property

Public Property Get SQL_StoredProcedures() As String
  SQL_StoredProcedures = _
    "sp_dboption sp_addmessage"
End Property

