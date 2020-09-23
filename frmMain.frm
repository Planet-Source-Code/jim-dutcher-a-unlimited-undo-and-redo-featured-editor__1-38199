VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Stuff2"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1680
      Top             =   3600
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4895
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":000C
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select &All"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   3975
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "&Paste"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdCut 
      Caption         =   "Cu&t"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdRedo 
      Caption         =   "&Redo"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                          Updated Code by Jim Dutcher                       '
'                 Original Code was created by Jason Shimkoski               '
'                                 Copyright 2000                             '
'                                                                            '
'        You can use this code in your apps, as long as you mention          '
'        me a in your About Box.                                             '
'                                                                            '
'        Please note that sections of this sample app were revisions of code '
'        found at www.visualstatement.com/vb                                 '
'                 http://falconsoft1.tripod.com                              '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'This updated version fixes the bug that caused the program to crash after
'the maximum number of undos are reached, this code will start overwritting
'the array that stores the undos once the specified max number of undos is reached

Option Explicit
'These are the variables for Undo and Redo

Const maxUndo = 50 'Maximum num of undos

Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(maxUndo) As String
Dim stackBK(maxUndo) As String
Dim i As Integer

'Here I'm just setting up the form.
Private Sub Form_Load()
    Form_Resize
End Sub

'Still just setting up the form
Private Sub Form_Resize()
    Me.Height = 4830
    Me.Width = 4305
End Sub



Private Sub cmdRedo_Click()
    'This is the basic redo stuff.
    If gintIndex < maxUndo Then ' max undo level is reached, do not redo
        gblnIgnoreChange = True
        gintIndex = gintIndex + 1
        On Error Resume Next
        rtfText.TextRTF = gstrStack(gintIndex)
        gblnIgnoreChange = False
    End If
End Sub

Private Sub cmdUndo_Click()
    'This says that if the Index is = to 0, then It shouldn't undo anymore
    If gintIndex = 0 Then Exit Sub
    
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    rtfText.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub rtfText_Change()
'Counter varibles, names have no meaning
Dim g As Integer
Dim b As Integer
Dim i As Integer

g = maxUndo 'Initialize this to the max number of undos

    If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        
        If gintIndex >= maxUndo + 1 Then 'If > max num of undos reached
        
            For b = 0 To maxUndo 'Copy the undo info to a backup array
                stackBK(b) = gstrStack(b)
            Next b
            
            For i = 0 To maxUndo 'Copy the backup array info back to the original, but in a different order
                If g >= 1 Then
                g = g - 1
                gstrStack(g) = stackBK(g + 1) 'gstrstack(49) = stackBK(50) get it??
                End If
            Next i
            
            gintIndex = maxUndo 'Set it to the max number of undos
            
        End If
        gstrStack(gintIndex) = rtfText.TextRTF
    End If
End Sub

Private Sub cmdCut_Click()
    'Clears the Clipboard to put text on it
    Clipboard.Clear
    'Sets the Text from rtfText onto the Clipboard
    Clipboard.SetText rtfText.SelText
    'Deletes the Selected Text on rtfText
    rtfText.SelText = ""
    'Sets the Focus to rtfText
    rtfText.SetFocus
End Sub

Private Sub cmdCopy_Click()
    'Clears the Clipboard to put text on it
    Clipboard.Clear
    'Sets the Text from rtfText onto the Clipboard
    Clipboard.SetText rtfText.SelText
    'Sets the Focus to rtfText
    rtfText.SetFocus
End Sub

Private Sub cmdPaste_Click()
    'Puts the Text from the clipboard into rtfText
    rtfText.SelText = Clipboard.GetText
    'Sets the Focus to rtfText
    rtfText.SetFocus
End Sub

Private Sub cmdSelectAll_Click()
    'Sets the cursors position to zero
    rtfText.SelStart = 0
    'Selects the full length of rtfText
    rtfText.SelLength = Len(rtfText.Text)
    'Sets the Focus to rtfText
    rtfText.SetFocus
End Sub

Private Sub Timer1_Timer()
Label1.Caption = gintIndex
End Sub
