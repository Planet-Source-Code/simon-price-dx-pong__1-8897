Attribute VB_Name = "DXpongMAIN"
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Type tBall
  x As Integer
  y As Integer
  xm As Integer
  ym As Integer
  Frame As Byte
  RollNow As Byte
End Type

Public Type tPaddle
  y As Integer
  ym As Integer
  Score As Integer
  LastMove As Byte
End Type

Dim Ball As tBall
Dim Bat(1 To 2) As tPaddle
Public i As Integer
Public Key As Byte
Public Frames As Integer

Const BATSPEED = 5

Const NOWT = 0
Const UP = 1
Const DOWN = 2


Sub Main()
Dim Cur As Long
Randomize Timer
MsgBox "Welcome to DX pong by Simon Price. The aim of the game is simply to stop the ball from passing you. Use the arrow keys to move up and down. You control the left hand side paddle. The computer will control the right hand side paddle. When you've had enough, press Escape to finish the game and be presented with some stats. Good Luck!", vbInformation, "DX Pong! by Simon Price"
'play sound while loading
sndPlaySound App.Path & "\SoundFX\Intro.wav", &H1
'hide the cursor
Cur = ShowCursor(0)
'start up DirectX
Form1.CrankItUp
'change screen res
ModDX7.SetDisplayMode 640, 480, 16
'load the graphix
ModSurfaces.LoadAllPics

'init the game variables
SetUpGame
'enter main game loop
MainGameLoop

'play sound while unloading
sndPlaySound App.Path & "\SoundFX\Exit.wav", &H1
'show cursor again
If Cur Then ShowCursor Cur Else ShowCursor 1
'change screen res back
ModDX7.RestoreDisplayMode
'close DirectX
Form1.ShutItDown
'show the results of the game
Results
MsgBox "Thankyou for playing DX pong by Simon Price. Please visit my website - www.VBgames.co.uk - for more cool VB games!", vbInformation, "Thankyou for playing DX pong - now visit my website!"
End
End Sub

Sub Results()
Select Case Bat(1).Score
Case 0
  Select Case Bat(2).Score
    Case 0
      MsgBox "Leaving already? But no-one has even scored yet!", vbInformation, "Results"
    Case Else
      MsgBox "Leaving already? That's probably because the computer scored " & Bat(2).Score & " and you didn't score any at all!", vbInformation, "Results"
  End Select
Case Is > Bat(2).Score
  MsgBox "Well done! You beat the computer by " & Bat(1).Score - Bat(2).Score & "points. You scored " & Bat(1).Score & " and the computer scored " & Bat(2).Score & " .", vbInformation, "Results"
Case Is < Bat(2).Score
  MsgBox "Unlucky, you lost to the computer by " & Bat(2).Score - Bat(1).Score & "points. You scored " & Bat(1).Score & " and the computer scored " & Bat(2).Score & " .", vbInformation, "Results"
Case Bat(2).Score
  MsgBox "What a match! You drew with the computer - both of you scored " & Bat(1).Score & " points!", vbInformation, "Results"
End Select
End Sub

Sub SetUpGame()

'place the ball
Ball.x = 320
Ball.y = 240
TryAgain:
Ball.xm = Int(Rnd * 20) - 10
If Abs(Ball.xm) < 5 Then GoTo TryAgain
Ball.ym = Int(Rnd * 10) - 5

'place the bats
For i = 1 To 2
    Bat(i).y = 240
    Bat(i).ym = 0
Next

End Sub

Sub MainGameLoop()
StartAgain:
On Error GoTo SortOutProbs
Dim x As Integer
Dim y As Integer

Do
DoEvents

Frames = Frames + 1
'Debug.Print Frames

If Key = vbKeyEscape Then Exit Sub

'draw background
ModDX7.SetRect SrcRect, 0, 0, 320, 240
ModDX7.SetRect DestRect, 0, 0, 640, 480
BackBuffer.Blt DestRect, Table, SrcRect, DDBLT_WAIT

''''''''''''''''''''''''''''''''''''''''''''
'move the paddles

'move paddle 1 according to keys

Bat(1).LastMove = NOWT

If Key = vbKeyUp Then
  MoveBat 1, UP
End If

If Key = vbKeyDown Then
  MoveBat 1, DOWN
End If

Bat(2).LastMove = NOWT

'move paddle 2 to according to ball position and velocity
If Ball.xm Then
  Select Case Ball.y + (((590 - Ball.x) / Ball.xm) * Ball.ym)
  
  Case -960 To Bat(2).y - BATSPEED - 960
    MoveBat 2, UP
  Case Bat(2).y + BATSPEED - 960 To -480
    MoveBat 2, DOWN
  
  Case -480 To Bat(2).y - BATSPEED - 480
    MoveBat 2, DOWN
  Case Bat(2).y + BATSPEED - 480 To 0
    MoveBat 2, UP

  Case 0 To Bat(2).y - BATSPEED
    MoveBat 2, UP
  Case Bat(2).y + BATSPEED To 480
    MoveBat 2, DOWN
    
  Case 480 To Bat(2).y - BATSPEED + 480
    MoveBat 2, DOWN
  Case Bat(2).y + BATSPEED + 480 To 960
    MoveBat 2, UP
  
  Case 960 To Bat(2).y - BATSPEED + 960
    MoveBat 2, UP
  Case Bat(2).y + BATSPEED + 960 To 1440
    MoveBat 2, DOWN
    
  End Select
End If

'draw paddles
ModDX7.SetRect SrcRect, 0, 0, 25, 100
BackBuffer.BltFast 25, Bat(1).y - 50, Paddles, SrcRect, DDBLTFAST_WAIT
ModDX7.SetRect SrcRect, 25, 0, 25, 100
BackBuffer.BltFast 590, Bat(2).y - 50, Paddles, SrcRect, DDBLTFAST_WAIT


''''''''''''''''''''''''''''''''''''''''''''
'move the balls

If Ball.RollNow = 3 Then

'spin the ball by changing frame
  Select Case Ball.xm
  Case Is > 0
    If Ball.Frame = 11 Then
    Ball.Frame = 0
    Else
    Ball.Frame = Ball.Frame + 1
    End If
  Case Is < 0
    If Ball.Frame = 0 Then
    Ball.Frame = 11
    Else
    Ball.Frame = Ball.Frame - 1
    End If
  Case Else
  
  End Select

  Ball.RollNow = 0

Else

  Ball.RollNow = Ball.RollNow + 1

End If

'move ball in x-direction
x = Ball.x + Ball.xm
  Select Case x
    
    Case 75 To 565
      Ball.x = x
    
    Case 25 To 75
      If Ball.xm < 0 Then
        Select Case Bat(1).y
          Case Ball.y - 50 To Ball.y + 50
            BatRebound 1, NOWT
          Case Ball.y - 70 To Ball.y - 50
            BatRebound 2, UP
          Case Ball.y + 50 To Ball.y + 70
            BatRebound 2, DOWN
          Case Else
            Ball.x = x
        End Select
      Else
        Ball.x = x
      End If
        
    Case 565 To 615
      
      If Ball.xm Then
        Select Case Bat(2).y
          Case Ball.y - 50 To Ball.y + 50
            BatRebound 2, NOWT
          Case Ball.y - 70 To Ball.y - 50
            BatRebound 2, UP
          Case Ball.y + 50 To Ball.y + 70
            BatRebound 2, DOWN
          Case Else
            Ball.x = x
        End Select
      Else
        Ball.x = x
      End If
    
    Case Is < 25
      HumanScore
    
    Case Is > 615
      ComputerScore
  End Select

'move ball in y-direction
  y = Ball.y + Ball.ym
  Select Case y
    Case 25 To 455
      Ball.y = y
    Case Else
      Ball.ym = -Ball.ym
      sndPlaySound App.Path & "\SoundFX\Clang.wav", &H1
  End Select
  
'draw balls
ModDX7.SetRect SrcRect, Ball.Frame * 50, 0, 50, 50
BackBuffer.BltFast Ball.x - 25, Ball.y - 25, Balls, SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
  

'flip into view
View.Flip Nothing, DDFLIP_WAIT

Loop

SortOutProbs:

ModDX7.WaitTillOK
GoTo StartAgain

End Sub

Sub BatRebound(BatNo As Byte, Spin As Byte)
'play boing sound
If BatNo = 1 Then
  sndPlaySound App.Path & "\SoundFX\Hit1.wav", &H1
Else
  sndPlaySound App.Path & "\SoundFX\Hit2.wav", &H1
End If

'speed up ball
Ball.xm = Ball.xm * -1.1
'now see if the ball was hit at an angle
Select Case Spin
Case NOWT
Ball.ym = Ball.ym * 1.1
Case DOWN
Ball.ym = Ball.ym * 1.1 - 5
Case UP
Ball.ym = Ball.ym * 1.1 + 5
End Select

'put spin on the ball
Select Case Bat(BatNo).LastMove
Case NOWT
  Ball.ym = Ball.ym + Int(Rnd * 4) - 2
Case UP
  Ball.ym = Ball.ym - Int(Rnd * 4)
Case DOWN
  Ball.ym = Ball.ym + Int(Rnd * 4)
End Select
End Sub

Sub MoveBat(BatNo As Byte, Move As Byte)

Bat(BatNo).LastMove = Move

Select Case Move
  Case UP
    If Bat(BatNo).y > 55 Then Bat(BatNo).y = Bat(BatNo).y - BATSPEED
  Case DOWN
    If Bat(BatNo).y < 425 Then Bat(BatNo).y = Bat(BatNo).y + BATSPEED
End Select

End Sub

Sub HumanScore()
Bat(2).Score = Bat(2).Score + 1
SetRect SrcRect, 0, 30, 200, 30
BackBuffer.BltFast 220, 225, Phrases, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
View.Flip Nothing, DDFLIP_WAIT
sndPlaySound App.Path & "\SoundFX\Lose.wav", 0
SetUpGame
End Sub

Sub ComputerScore()
Bat(1).Score = Bat(1).Score + 1
SetRect SrcRect, 0, 0, 200, 30
BackBuffer.BltFast 220, 225, Phrases, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
View.Flip Nothing, DDFLIP_WAIT
sndPlaySound App.Path & "\SoundFX\Win.wav", 0
SetUpGame
End Sub
