VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Undo Test"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5055
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdRedoAll 
      Caption         =   "Redo All"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdUndoAll 
      Caption         =   "Undo All"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdResetUndo 
      Caption         =   "Clear Undo Stack"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CheckBox checkIgnoreChanges 
      Caption         =   "Ignore Changes"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      ToolTipText     =   "Turns on/off the tracking of text changes"
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdCut 
      Caption         =   "Cut"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdRedo 
      Caption         =   "Redo"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   615
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2778
      _Version        =   393217
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.Label Label2 
      Caption         =   "(c) 2002 by Sebastian Thomschke               http://www.sebthom.de"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Multiple Undos/Redos v1.03"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************
' Multi Undo Class for RichTextBox - Demonstration
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
Private WithEvents Undo As clsUndo
Attribute Undo.VB_VarHelpID = -1

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const SurpriseText = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 MS Sans Serif;}}" + _
"\viewkind4\uc1\pard\f0\fs17 This is a Test. \par {\pict\wmetafile8\picw624\pich624\picwgoal354\pichgoal354 " + _
"0100090000037403000000005e0300000000050000000b0200000000050000000c02700270025e030000430f2000cc00000017001700000000007002700200000000280000001700000017000000010018000000000078060000130b0000130b00000000000000000000fffffffffffffffffffffffffffffffffffffffffffffbfbfee6e8fdc7cbfca1a7fc9ea4fca1a7fdc9cdfeecedfffefeffff" + _
"ffffffffffffffffffffffffffffffffffffff000000fffffffffffffffffffffffffffffffff9f9fcacb2f9515df72130f60b1bf60011f60011f60011f60d1df82b39f9636dfdc6cafffefeffffffffffffffffffffffffffffff000000fffffffffffffffffffffffffed5d8fa6670f60a1af60011f60011f60011f60011f60011f60011f60011f60011f60011f71827fa7881fee7e9ffffffffff" + _
"ffffffffffffff000000fffffffffffffffefefc9fa5f72231f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60112f83946fdc3c7ffffffffffffffffff000000fffffffffffffdc6caf71827f60011f60011f60011f60011f60011f60011f60415f60d1df60415f60011f60011f60011f60011f60011f60011f83946fee7e9ffffffffffff000000ffff" + _
"fffff1f2f94753f60011f60011f60011f60011f71121f95964fb939afcb2b7fdd1d4fcb2b7fb939af94b57f60617f60011f60011f60011f60112fb878ffffefeffffff000000fffffffb99a0f60617f60011f60011f60011f82d3bfcb5bafffcfcfffffffffffffffffffffffffffffffff5f6fb939af71827f60011f60011f60011f71827fdc6caffffff000000feeeeff83240f60011f60011f600" + _
"11f72a38fed6d9fff9f9fdc3c7fc9ea4fc9ea4fc9ea4fc9ea4fca1a7fdcdd0fff8f8fca3a9f60617f60011f60011f60011f9636dfffefe000000fdd3d6f60e1ef60011f60011f60112fb8e96fdc8ccf95762f60e1ef60011f60011f60011f60011f60011f71121f95a65fdced1f8404df60011f60011f60011f82b39feeced000000fcb2b7f60415f60011f60011f72130fcb5baf72634f60011f600" + _
"11f60011f60011f60011f60011f60011f60011f60011f94a56fb959cf60415f60011f60011f60d1dfdc9cd000000fb939af60011f60011f60011f83341f8424ff60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60314f95863f71827f60011f60011f60011fb969d000000fa656ff60011f60011f60011f71525f60617f60011f60011f60011f60011f60011f60011f600" + _
"11f60011f60011f60011f60011f71020f60c1cf60011f60011f60011fa6973000000fb878ff60011f60011f60011f60011f60011f60011f60112f60415f60011f60011f60011f60011f60011f60818f60112f60011f60011f60011f60011f60011f60011fb969d000000fca8aef60112f60011f60011f60011f60011f60a1afb868efca0a6f71f2ef60011f60011f60011f83c49fcb6bbfa767ff601" + _
"12f60011f60011f60011f60011f60b1bfdc7cb000000fdced1f60d1df60011f60011f60011f60011f94854fff9f9fffffffa7982f60011f60011f60818fcb6bbfffffffedcdef71b2af60011f60011f60011f60011f72130fee6e8000000fee6e8f72433f60011f60011f60011f60011f72a38fdc6cafed5d8f83e4bf60011f60011f60112fb8088fee4e6fc9da4f60617f60011f60011f60011f600" + _
"11f9515dfffbfb000000fffefefa7881f60112f60011f60011f60011f60011f71121f71827f60112f60011f60011f60011f60415f7202ff60e1ef60011f60011f60011f60011f60a1afcacb2ffffff000000fffffffee4e6f82f3df60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011fa6670fff9f9ffffff000000ffff" + _
"fffffffffca9aff60a1af60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f72231fed5d8ffffffffffff000000fffffffffffffff9f9fa727bf60a1af60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f60011f71827fc9fa5ffffffffffffffffff000000fffffffffffffffffffff9f9fcac" + _
"b2f83946f60112f60011f60011f60011f60011f60011f60011f60011f60011f60011f60617f94753fdc6cafffefeffffffffffffffffff000000fffffffffffffffffffffffffffffffee7e9fa7881f82b39f60e1ef60415f60011f60011f60011f60415f60e1ef83240fb99a0fff1f2ffffffffffffffffffffffffffffff000000fffffffffffffffffffffffffffffffffffffffefefeecedfdd3" + _
"d6fcb2b7fb939afa6973fb939afcb2b7fdd3d6feeeefffffffffffffffffffffffffffffffffffffffffff000000030000000000" + _
"} \par\b \uld Surprise! Surprise! \ul0 Even undoing of image insertions and font formatting is supported." + _
"\par }"


'*********************************************************************
' Form Events
'*********************************************************************
Private Sub Form_Load()
   Set Undo = New clsUndo
   Undo.Create Controls, RichTextBox1
   RichTextBox1.SelText = "This"
   Undo.TakeUndoSnapShot
   RichTextBox1.SelText = " is"
   Undo.TakeUndoSnapShot
   RichTextBox1.SelText = " a"
   Undo.TakeUndoSnapShot
   RichTextBox1.SelText = " Test."
   Undo.TakeUndoSnapShot
   RichTextBox1.TextRTF = SurpriseText
   Undo.TakeUndoSnapShot
   Undo.Undo
   Call RichTextBox1_SelChange
End Sub


'*********************************************************************
' RichtextBox1 Events
'*********************************************************************
Private Sub RichTextBox1_SelChange()
   cmdCut.Enabled = RichTextBox1.SelLength
   cmdCopy.Enabled = RichTextBox1.SelLength
   cmdPaste.Enabled = Not IsClipboardEmpty
End Sub


'*********************************************************************
' Undo Events
'*********************************************************************
Private Sub Undo_StateChanged()
   cmdUndo.Enabled = Undo.canUndo
   cmdRedo.Enabled = Undo.canRedo
   cmdUndoAll.Enabled = Undo.canUndo
   cmdRedoAll.Enabled = Undo.canRedo
End Sub


'*********************************************************************
' Button Events
'*********************************************************************
Private Sub cmdCut_Click()
   SendMessage RichTextBox1.hWnd, WM_CUT, 0, 0
   RichTextBox1.SetFocus
End Sub

Private Sub cmdCopy_Click()
   SendMessage RichTextBox1.hWnd, WM_COPY, 0, 0
   RichTextBox1.SetFocus
End Sub

Private Sub cmdPaste_Click()
   On Error Resume Next
   SendMessage RichTextBox1.hWnd, WM_PASTE, 0, 0
   RichTextBox1.SetFocus
End Sub

Private Sub checkIgnoreChanges_Click()
   Undo.IgnoreChange = checkIgnoreChanges.Value
End Sub

Private Sub cmdUndo_Click()
   Undo.Undo
   RichTextBox1.SetFocus
End Sub

Private Sub cmdRedo_Click()
   Undo.Redo
   RichTextBox1.SetFocus
End Sub

Private Sub cmdRedoAll_Click()
   Undo.Redo Undo.getRedoCount
End Sub

Private Sub cmdUndoAll_Click()
   Undo.Undo Undo.getUndoCount
End Sub

Private Sub cmdResetUndo_Click()
   Undo.Reset
End Sub


'*********************************************************************
' other methods
'*********************************************************************
Private Function IsClipboardEmpty() As Boolean
   On Error Resume Next
   IsClipboardEmpty = True
   If Clipboard.GetFormat(vbCFText) Then IsClipboardEmpty = False
   If Clipboard.GetFormat(vbCFBitmap) Then IsClipboardEmpty = False
   If Clipboard.GetFormat(vbCFDIB) Then IsClipboardEmpty = False
   If Clipboard.GetFormat(vbCFRTF) Then IsClipboardEmpty = False
End Function
