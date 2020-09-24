VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUndoElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' changed area before/after editing
Public TextBlock_BeforeChange As String
Public TextBlock_AfterChange As String

' number of characters that didn't change
Public MatchingCharsFromLeft As Long
Public MatchingCharsFromRight As Long

' for cursor/selection
Public SelStart_BeforeChange As Long
Public SelLength_BeforeChange As Long
Public SelStart_AfterChange As Long
Public SelLength_AfterChange As Long
