VERSION 5.00
Begin VB.Form frmBoard 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BitBlt API Example"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   Icon            =   "frmBoard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Information"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop!"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.HScrollBar hsSpeed 
      Height          =   255
      Left            =   120
      Max             =   20
      Min             =   1
      TabIndex        =   2
      Top             =   5160
      Value           =   6
      Width           =   2535
   End
   Begin VB.PictureBox picBoard 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   0
      ScaleHeight     =   317
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   485
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.PictureBox picBall 
         AutoRedraw      =   -1  'True
         Height          =   525
         Left            =   5520
         Picture         =   "frmBoard.frx":000C
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   1
         Top             =   2160
         Width           =   525
      End
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Speed Control:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   1050
   End
End
Attribute VB_Name = "frmBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim Running As Boolean, Rite As Boolean, Down As Boolean

Private Sub Pause(seconds As Integer)
    'This sub lets you pause the application while allowing things to keep working.
    Dim Start As Single, Finish As Single
    Start = Timer                    'Start the timer event.
    Do While Timer < Start + seconds 'If it still hasn't been <seconds> seconds... keep pausing.
        DoEvents                     'Let the processor do it's business. **NEEDED**
    Loop
    Finish = Timer                   'Clear the timer event.
End Sub

Private Sub MoveBall()
    'This routine moves the ball around.
    Dim y As Long, x As Long 'The y and x axis variables.
    y = 0                    'Start off with zero on the y axis.
    x = 0                    'Start off with zero on the x axis.
    Down = True              'Start off with moving down on y axis.
    Rite = True              'Start off with moving right on y axis.
    Do While Running = True  'If the user hasn't pressed the Stop button, keep going.
        picBoard.Cls         'Clear old graphics for repaint.
        Sleep hsSpeed.Value  'Sleep for 1 to 20 milliseconds, depending on the scrollbar value.
        BitBlt picBoard.hDC, x, y, picBall.ScaleWidth, picBall.ScaleHeight, picBall.hDC, picBall.ScaleLeft, picBall.ScaleTop, vbSrcCopy
        If Down = True Then y = y + 1 Else: y = y - 1
        If Rite = True Then x = x + 1 Else: x = x - 1
        If x >= picBoard.ScaleWidth - picBall.ScaleWidth Or x <= 0 Then Rite = Not Rite
        If y >= picBoard.ScaleHeight - picBall.ScaleHeight Or y <= 0 Then Down = Not Down
        DoEvents             'Let the processor do its important stuff. **NEEDED**
    Loop
End Sub

Private Sub cmdInfo_Click()
MsgBox "This example of BitBlt shows you how to move an image around in an object with a DC (device context). Many people like BitBlt because it's fast and flicker free. Using BitBlt is a great way of  creating 2D games."
End Sub

Private Sub cmdStart_Click()
    Running = True
    cmdStop.Enabled = True
    cmdStart.Enabled = False
    Call MoveBall
End Sub

Private Sub cmdStop_Click()
    Running = False 'Set the running condition to false.
    cmdStart.Enabled = True
    cmdStop.Enabled = False
End Sub

Private Sub Form_Load()
    picBoard.AutoRedraw = True   'Force AutoRedraw.
    picBall.Left = Me.Width + 10 'Hide the original ball, since we paint copies.
    Me.Show                      'Force the screen to load.
    Pause 1                      'Let the form load completely before moving on.
    Running = True               'Start the running condition.
    Call MoveBall                'Call the routine to move the ball.
End Sub

Private Sub hsSpeed_Change()
    lblSpeed.Caption = hsSpeed.Value 'Display the value.
End Sub
