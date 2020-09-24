VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUndo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************
' Multi Undo Class for RichTextBox v1.03
' Copyright ©2002 by Sebastian Thomschke, All Rights Reserved.
' http://www.sebthom.de
'*********************************************************************
' If you like this code, please vote for it at Planet-Source-Code.com:
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=34094
' Thank you
'*********************************************************************
' You are free to use this code within your own applications, but you
' are expressly forbidden from selling or otherwise distributing this
' source code without prior written consent.
'*********************************************************************
' Thanks to MrBobo for his suggestions for improvement
'*********************************************************************
Option Explicit


'*********************************************************************
' API Declarations
'*********************************************************************
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function InvalidateClientRect Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long

Private Const WM_SETREDRAW = &HB


'*********************************************************************
' Private Vars
'*********************************************************************
Private WithEvents MyRichTextBox As RichTextBox
Attribute MyRichTextBox.VB_VarHelpID = -1
Private WithEvents MyTimer As VB.Timer
Attribute MyTimer.VB_VarHelpID = -1

Private TrackUndo As Boolean
Private UndoStack As Collection
Private RedoStack As Collection

Private SelStart_BeforeChange As Long
Private SelLength_BeforeChange As Long
Private SelStart_AfterChange As Long
Private SelLength_AfterChange As Long

Private TextRTF_BeforeChange As String


'*********************************************************************
' Public Events
'*********************************************************************
Public Event StateChanged()


'*********************************************************************
' Public Properties
'*********************************************************************
Public Property Get canUndo() As Boolean
   canUndo = UndoStack.Count > 1 And TrackUndo
End Property

Public Property Get canRedo() As Boolean
   canRedo = RedoStack.Count > 0 And TrackUndo
End Property

Public Property Let IgnoreChange(ByVal IIgnoreChange As Boolean)
   ' turns the tracking of changes on or off
   
   ' when tracking is enabled again, we create an undo item to not get out of sync
   ' TakeUndoSnapShot is intelligent enough not to create an undoitem if no changes took part since the turn off
   If Not IIgnoreChange And Not TrackUndo Then TakeUndoSnapShot
   
   TrackUndo = Not IIgnoreChange

   RaiseEvent StateChanged
End Property

Public Property Get IgnoreChange() As Boolean
   IgnoreChange = Not TrackUndo
End Property

Public Property Get getUndoCount() As Long
   getUndoCount = UndoStack.Count - 1
End Property

Public Property Get getRedoCount() As Long
   getRedoCount = RedoStack.Count
End Property


'*********************************************************************
' Public Methods
'*********************************************************************
Public Function Create(ByRef Controls As Object, ByRef RichTextBox As RichTextBox, Optional Delay As Long = 250) As Boolean

   On Error GoTo Create_Error
   
   Set MyRichTextBox = RichTextBox
   
   ' the timer is used to avoid manual undoing of every single character
   ' instead inputs made within a defined time are grouped to one undo action
   Set MyTimer = Controls.Add("VB.Timer", "UndoTimer_" & MyRichTextBox.Name)
   MyTimer.Enabled = False
   MyTimer.Interval = Delay
   
   Reset
   
   Create = True
   Exit Function
   
Create_Error:
   MsgBox _
      Err.Description, _
      vbOKOnly + vbExclamation + vbMsgBoxHelpButton, _
      Err.Source & " - Unexpected Error No. " & Err.Number, _
      Err.HelpFile, _
      Err.HelpContext
      
   Create = False
   Exit Function
End Function

Public Sub Reset()
   Set UndoStack = New Collection
   Set RedoStack = New Collection
   
   TextRTF_BeforeChange = MyRichTextBox.TextRTF
   
   Call MyRichTextBox_SelChange
      
   TrackUndo = True
   
   TakeUndoSnapShot
   
   RaiseEvent StateChanged
End Sub

Public Sub Redo(Optional ByVal RedoSteps As Long = 1)
   If RedoStack.Count < 1 Or RedoSteps < 1 Then Exit Sub
   
   TrackUndo = False
   
   If RedoSteps > RedoStack.Count Then RedoSteps = RedoStack.Count
   
   Dim RedoItem As clsUndoElement, i As Long
   
   For i = 1 To RedoSteps
   
      Set RedoItem = RedoStack.Item(RedoStack.Count)

      ' redo the text changes
      MyRichTextBox.TextRTF = Left$(MyRichTextBox.TextRTF, RedoItem.MatchingCharsFromLeft) & RedoItem.TextBlock_AfterChange & Right$(MyRichTextBox.TextRTF, RedoItem.MatchingCharsFromRight)
      TextRTF_BeforeChange = MyRichTextBox.TextRTF
   
      ' set the selection
      MyRichTextBox.SelStart = RedoItem.SelStart_AfterChange
      MyRichTextBox.SelLength = RedoItem.SelLength_AfterChange
   
      ' Moving the undo item to the undostack
      UndoStack.Add Item:=RedoItem
      RedoStack.Remove RedoStack.Count
      
   Next
   TrackUndo = True
   
   RaiseEvent StateChanged
End Sub

Public Sub Undo(Optional ByVal UndoSteps As Long = 1)
   If UndoStack.Count < 2 Or UndoSteps < 1 Then Exit Sub
   
   TrackUndo = False
   
   If UndoSteps >= UndoStack.Count Then UndoSteps = UndoStack.Count - 1
   
   Dim UndoItem As clsUndoElement, i As Long
   
   For i = 1 To UndoSteps
   
      Set UndoItem = UndoStack.Item(UndoStack.Count)

      SetRefresh MyRichTextBox.hWnd, False
   
      ' undo the text changes
      MyRichTextBox.TextRTF = Left$(MyRichTextBox.TextRTF, UndoItem.MatchingCharsFromLeft) & UndoItem.TextBlock_BeforeChange & Right$(MyRichTextBox.TextRTF, UndoItem.MatchingCharsFromRight)
      MyRichTextBox.SelStart = UndoItem.SelStart_BeforeChange
      MyRichTextBox.SelLength = UndoItem.SelLength_BeforeChange
      SetRefresh MyRichTextBox.hWnd, True
   
      TextRTF_BeforeChange = MyRichTextBox.TextRTF
   
      ' Moving the undo item to the redostack
      RedoStack.Add Item:=UndoItem
      UndoStack.Remove UndoStack.Count
      
   Next
   
   TrackUndo = True
   
   RaiseEvent StateChanged
End Sub

Public Sub TakeUndoSnapShot()
   ' add the new undo element
   Dim UndoItem As New clsUndoElement, _
       TextRTF_AfterChange As String, _
       Len_TextRTF_AfterChange As Long, _
       Len_TextRTF_BeforeChange As Long
   
   TextRTF_AfterChange = MyRichTextBox.TextRTF
   
   Len_TextRTF_AfterChange = Len(TextRTF_AfterChange)
   Len_TextRTF_BeforeChange = Len(TextRTF_BeforeChange)
   
   UndoItem.MatchingCharsFromLeft = getMatchingCharCount( _
      TextRTF_BeforeChange, _
      TextRTF_AfterChange _
   )
   
   UndoItem.MatchingCharsFromRight = getMatchingCharCount( _
      Right$(TextRTF_BeforeChange, Len_TextRTF_BeforeChange - UndoItem.MatchingCharsFromLeft), _
      Right$(TextRTF_AfterChange, Len_TextRTF_AfterChange - UndoItem.MatchingCharsFromLeft), _
      True _
   )
   
   UndoItem.TextBlock_BeforeChange = Mid$( _
      TextRTF_BeforeChange, _
      UndoItem.MatchingCharsFromLeft + 1, _
      Len_TextRTF_BeforeChange - UndoItem.MatchingCharsFromLeft - UndoItem.MatchingCharsFromRight _
   )
   
   UndoItem.TextBlock_AfterChange = Mid$( _
      TextRTF_AfterChange, _
      UndoItem.MatchingCharsFromLeft + 1, _
      Len_TextRTF_AfterChange - UndoItem.MatchingCharsFromLeft - UndoItem.MatchingCharsFromRight _
   )
      
   UndoItem.SelStart_BeforeChange = SelStart_BeforeChange
   UndoItem.SelLength_BeforeChange = SelLength_BeforeChange
   UndoItem.SelStart_AfterChange = MyRichTextBox.SelStart
   UndoItem.SelLength_AfterChange = MyRichTextBox.SelLength
      
   If UndoStack.Count > 0 Then
      'No change
      If Len_TextRTF_BeforeChange = Len_TextRTF_AfterChange And UndoItem.MatchingCharsFromLeft = Len_TextRTF_BeforeChange Then Exit Sub
   End If
   
   UndoStack.Add Item:=UndoItem
            
   ' remove redo elements
   Set RedoStack = New Collection
      
   TextRTF_BeforeChange = TextRTF_AfterChange
     
   SelStart_BeforeChange = MyRichTextBox.SelStart
   SelLength_BeforeChange = MyRichTextBox.SelLength
   SelStart_AfterChange = MyRichTextBox.SelStart
   SelLength_AfterChange = MyRichTextBox.SelLength
      
   RaiseEvent StateChanged
End Sub


'*********************************************************************
' MyRichTextBox Events
'*********************************************************************
Private Sub MyRichTextBox_Change()
   If TrackUndo Then
      ' restarting the timer by switching off and on
      MyTimer.Enabled = False
      MyTimer.Enabled = True
   Else
      MyTimer.Enabled = False
   End If
End Sub

Private Sub MyRichTextBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' a mouse click will break the delay and force an undosnapshot if text has changed
   If MyTimer.Enabled = True Then Call MyTimer_Timer
End Sub

Private Sub MyRichTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim CtrlOnly As Boolean

   CtrlOnly = (Shift = vbCtrlMask)
   
   'CTRL+Z = Undo
   If CtrlOnly And (KeyCode = vbKeyZ) Then
      
      Undo
      KeyCode = 0
      
   'CTRL+Y = Redo
   ElseIf CtrlOnly And (KeyCode = vbKeyY) Then
      
      Redo
      KeyCode = 0
      
   ' navigating the cursor and space will break the delay and force an undosnapshot if text has changed
   ElseIf _
      (KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12) Or _
      KeyCode = vbKeyReturn Or _
      (KeyCode >= vbKeySpace And KeyCode <= vbKeySnapshot) Then
      If MyTimer.Enabled = True Then Call MyTimer_Timer
   End If
End Sub

Private Sub MyRichTextBox_SelChange()
   If Not MyTimer.Enabled Then
      SelStart_BeforeChange = SelStart_AfterChange
      SelLength_BeforeChange = SelLength_AfterChange
      SelStart_AfterChange = MyRichTextBox.SelStart
      SelLength_AfterChange = MyRichTextBox.SelLength
   End If
End Sub


'*********************************************************************
' MyTimer Events
'*********************************************************************
Private Sub MyTimer_Timer()
   MyTimer.Enabled = False

   If TrackUndo Then TakeUndoSnapShot
End Sub


'*********************************************************************
' Private Methods
'*********************************************************************
Private Function getMatchingCharCount(ByVal StringA As String, ByVal StringB As String, Optional FromRightToLeft As Boolean = False) As Long
   ' compares two strings from the left to the right and returns the count of characters
   ' upto which both strings are identical
   ' for performance reasons a binary search like comparison algorithm is used
   ' (c) 2002 Sebastian Thomschke
   
   Dim Low As Long, High As Long, Middle As Long, isMatching As Boolean
   Dim MatchingLen As Long, LenStringA As Long, LenStringB As Long
   
   Low = 1
   High = IIf(Len(StringA) < Len(StringB), Len(StringA), Len(StringB))
   If High = 0 Then GoTo getMatchingCharCount_FoundIt
   
   LenStringA = Len(StringA)
   LenStringB = Len(StringB)
   
   getMatchingCharCount = 0
   Do
      Middle = Low + ((High - Low) \ 2)
      If FromRightToLeft Then
         isMatching = (Mid$(StringA, LenStringA - Middle + 1, Middle - MatchingLen) = Mid$(StringB, LenStringB - Middle + 1, Middle - MatchingLen))
      Else
         isMatching = (Mid$(StringA, MatchingLen + 1, Middle - MatchingLen) = Mid$(StringB, MatchingLen + 1, Middle - MatchingLen))
      End If
      If isMatching Then
         ' match is in upper half
         MatchingLen = Middle
         Low = Middle + 1
         If Low > High Then GoTo getMatchingCharCount_FoundIt
      Else
         ' match is in lower half
         High = Middle - 1
         If High = 0 Then   ' No match found
            GoTo getMatchingCharCount_FoundIt
         ElseIf Low > High Then
            GoTo getMatchingCharCount_FoundIt
         End If
      End If
   Loop
   Exit Function
   
getMatchingCharCount_FoundIt:
   getMatchingCharCount = High
   Exit Function
End Function

Private Sub SetRefresh(ByVal hWnd As Long, ByVal EnableRefresh As Boolean)
   If EnableRefresh Then
      SendMessage hWnd, WM_SETREDRAW, 1, 0
      InvalidateClientRect hWnd, 0, 1
   Else
      SendMessage hWnd, WM_SETREDRAW, 0, 0
   End If
End Sub
