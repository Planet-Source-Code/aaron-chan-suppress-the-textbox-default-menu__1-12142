VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suppress the Text-box Right-click menu"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtTest 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmTest.frx":000C
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuItems 
      Caption         =   "&Items"
      Visible         =   0   'False
      Begin VB.Menu mnuItem0 
         Caption         =   "Item 0"
      End
      Begin VB.Menu mnuItem1 
         Caption         =   "Item 1"
      End
      Begin VB.Menu mnuItem2 
         Caption         =   "Item 2"
      End
      Begin VB.Menu mnuItem3 
         Caption         =   "Item 3"
      End
      Begin VB.Menu mnuItem4 
         Caption         =   "Item 4"
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'lock updates api (pass 0 as param to stop)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'when user clicks Done button
Private Sub cmdDone_Click()
    End 'ends the program
End Sub

'when user clicks File | Exit
Private Sub mnuFileExit_Click()
    End 'ends the program
End Sub

'main part of the example
'note that this is the MouseUp event of txtTest
Private Sub txtTest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this checks if the user clicked the textbox
    'with the right button
    If Button = vbRightButton Then
        'if they did, disable textbox because
        'the textbox will not respond to mouse
        'events when disabled
        txtTest.Enabled = False
        
        'lock the window so that the user won't
        'see the textbox as being disabled
        LockWindowUpdate txtTest.hWnd
        
        PopupMenu mnuItems 'popup the menu
        
        UnlockWindowUpdate 'unlock textbox
        
        'enable textbox again, phew, you did it!
        txtTest.Enabled = True
    End If
End Sub

'unlock window update by passing 0 to the
'LockWindowUpdate api
Private Sub UnlockWindowUpdate()
    LockWindowUpdate 0
End Sub
