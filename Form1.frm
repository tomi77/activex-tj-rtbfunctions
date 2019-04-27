VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   4455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4683
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private r As TJ_RTBFunctions.Colouring

Private Sub Command1_Click()
  r.ColourCode RichTextBox1, True
End Sub

Private Sub Form_Load()
  Set r = New TJ_RTBFunctions.Colouring
  r.Initialize btWord, _
    Array( _
      r.SQL_DataTypes & r.Delimiter & r.SQL_KeyWords, _
      r.SQL_Functions & r.Delimiter & r.SQL_StoredProcedures, _
      r.SQL_Operators), _
    Array( _
      RGB(0, 0, 255), _
      RGB(), _
      RGB()), _
    r.Delimiter
  r.Initialize btFromTo, _
    Array( _
      r.SQL_CommentBlock, _
      r.SQL_Text), _
    Array( _
      RGB(), _
      RGB()), _
    r.Delimiter
  r.Initialize btFromToEOL, Array(), Array(), " "
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set r = Nothing
End Sub


