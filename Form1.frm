VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2508
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3744
   ForeColor       =   &H0000FF00&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":030A
   MousePointer    =   99  'Custom
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CrankItUp()
ModDX7.Init Me.hwnd
End Sub

Sub ShutItDown()
ModDX7.EndIt Me.hwnd
End Sub

Private Sub Form_Click()
ShutItDown
End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Key = KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Key = 0
End Sub

Private Sub Form_Load()
Move 0, 0, 640, 480
End Sub

Private Sub Timer1_Timer()
Debug.Print Frames
Frames = 0
End Sub
