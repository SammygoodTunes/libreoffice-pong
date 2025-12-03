REM *** libreoffice-pong by SammygoodTunes (2025) ***

Option Explicit

Global KeyHandler
Global PaddleUp
Global PaddleDown
Global PrevKey

Global FrameW
Global FrameH

Sub Main

FrameW = 78
FrameH = 50

Dim PaddleX As Integer
Dim PaddleY As Integer
Dim PaddleW As Integer
Dim PaddleH As Integer

Dim Paddle2X As Integer
Dim Paddle2Y As Integer
Dim Paddle2W As Integer
Dim Paddle2H As Integer

Dim BallX As Double
Dim BallY As Double
Dim BallPrevX As Double
Dim BallPrevY As Double
Dim BallVelocityX As Double
Dim BallVelocityY As Double
Dim BallAngle As Double

Dim PlayerScoreCell as Object
Dim OpponentScoreCell as Object
Dim PlayerScore as Integer
Dim OpponentScore as Integer

Dim ClearBallTrail as Boolean
Dim OpponentDelay as Double
Dim Diff As Double

Dim PongSheet As Object, Cell As Object

PlayerScore = 0
OpponentScore = 0

PaddleX = 0
PaddleY = 0
PaddleH = 6

Paddle2X = FrameW - 1
Paddle2Y = 0
Paddle2H = 6

OpponentDelay = (FrameW / 2.75 * Rnd)

ResetBall(BallX, BallY, BallAngle, BallVelocityX, BallVelocityY)

If BallVelocityX = 0 Then 
	BallVelocityX = -1
End If
If BallVelocityY = 0 Then 
	BallVelocityY = -1
End If

PaddleUp = False
PaddleDown = False

KeyHandler = CreateUnoListener("KeyHandler_", "com.sun.star.awt.XKeyHandler")
ThisComponent.CurrentController.addKeyHandler(KeyHandler)

If Not ThisComponent.Sheets.hasByName("PONG") Then
	ThisComponent.Sheets.insertNewByName("PONG", ThisComponent.Sheets.getCount)
End If

PongSheet = ThisComponent.Sheets.getByName("PONG")
PongSheet.getCellRangeByName("A1:BZ50").Rows.Height = 200
PongSheet.getCellRangeByName("A1:BZ50").Columns.Width = 200

ClearAndRedraw(PongSheet, PaddleX, PaddleY, PaddleH, Paddle2X, Paddle2Y, Paddle2H)

PlayerScoreCell = PongSheet.getCellRangeByName("CA51")
OpponentScoreCell = PongSheet.getCellRangeByName("CB51")
PlayerScoreCell.Columns.Width = 2700
PlayerScoreCell.String = "Player Score: " & PlayerScore
OpponentScoreCell.Columns.Width = 3200
OpponentScoreCell.String = "Opponent Score: " & OpponentScore

While True
	ClearBallTrail = True
	BallPrevX = BallX
	BallPrevY = BallY
	ThisComponent.getCurrentController().select(PongSheet.getCellByPosition(0, 0))

	If PaddleUp And PaddleY > 0 Then
		PaddleY = PaddleY - 1
		PongSheet.getCellByPosition(PaddleX, PaddleY).CellBackColor = RGB(0, 0, 0)
		PongSheet.getCellByPosition(PaddleX, PaddleY + PaddleH + 1).CellBackColor = RGB(255, 255, 255)
	End If
	
	If PaddleDown And PaddleY < FrameH - 1 - PaddleH Then
		PaddleY = PaddleY + 1
		PongSheet.getCellByPosition(PaddleX, PaddleY - 1).CellBackColor = RGB(255, 255, 255)
		PongSheet.getCellByPosition(PaddleX, PaddleY + PaddleH).CellBackColor = RGB(0, 0, 0)
	End If
	
	If Paddle2Y > 0 And BallY < Paddle2Y And BallX > FrameW \ 2 + OpponentDelay Then
		Paddle2Y = Paddle2Y - 1
		PongSheet.getCellByPosition(Paddle2X, Paddle2Y).CellBackColor = RGB(0, 0, 0)
		PongSheet.getCellByPosition(Paddle2X, Paddle2Y + Paddle2H + 1).CellBackColor = RGB(255, 255, 255)

	End If
	
	If Paddle2Y < FrameH - 1 - Paddle2H And BallY > Paddle2Y + Paddle2H And BallX > FrameW \ 2 + OpponentDelay Then
		Paddle2Y = Paddle2Y + 1
		PongSheet.getCellByPosition(Paddle2X, Paddle2Y - 1).CellBackColor = RGB(255, 255, 255)
		PongSheet.getCellByPosition(Paddle2X, Paddle2Y + Paddle2H).CellBackColor = RGB(0, 0, 0)
	End If
	
	
	If Int(BallX) < 1 Then
		If Int(BallY) < PaddleY Or Int(BallY) > PaddleY + PaddleH Then
			PongSheet.getCellByPosition(Int(BallX), Int(BallY)).CellBackColor = RGB(255, 255, 255)
			OpponentScore = OpponentScore + 1
			ResetBall(BallX, BallY, BallAngle, BallVelocityX, BallVelocityY)
			PongSheet.getCellRangeByName("CB51").String = "Opponent Score: " & OpponentScore
			If BallVelocityX < 0 Then
				BallVelocityX = -BallVelocityX
			EndIf
			ClearAndRedraw(PongSheet, PaddleX, PaddleY, PaddleH, Paddle2X, Paddle2Y, Paddle2H)
		Else
			Diff = ((PaddleY + PaddleH / 2) - BallY) / (PaddleH / 2)
			ClearBallTrail = False
			PongSheet.getCellByPosition(Int(BallX), Int(BallY)).CellBackColor = RGB(0, 0, 0)
			BallAngle = (5 * 3.141592 / 12) * Diff
			UpdateBallVelocity(BallAngle, BallVelocityX, BallVelocityY)
			If BallVelocityX < 0 Then
				BallVelocityX = -BallVelocityX
			EndIf
			OpponentDelay = (FrameW / 2.75 * Rnd)
		End If
	End If
	
	If Int(BallX) > FrameW - 2 Then
		If Int(BallY) < Paddle2Y Or Int(BallY) > Paddle2Y + Paddle2H Then
			PongSheet.getCellByPosition(Int(BallX), Int(BallY)).CellBackColor = RGB(255, 255, 255)
			PlayerScore = PlayerScore + 1
			ResetBall(BallX, BallY, BallAngle, BallVelocityX, BallVelocityY)
			PongSheet.getCellRangeByName("CA51").String = "Player Score: " & PlayerScore
			If BallVelocityX > 0 Then
				BallVelocityX = -BallVelocityX
			EndIf
			ClearAndRedraw(PongSheet, PaddleX, PaddleY, PaddleH, Paddle2X, Paddle2Y, Paddle2H)
		Else
			Diff = ((Paddle2Y + Paddle2H / 2) - BallY) / (Paddle2H / 2)
			ClearBallTrail = False
			PongSheet.getCellByPosition(Int(BallX), Int(BallY)).CellBackColor = RGB(0, 0, 0)
			BallAngle = (5 * 3.141592 / 12) * Diff
			UpdateBallVelocity(BallAngle, BallVelocityX, BallVelocityY)
			BallVelocityX = -BallVelocityX
		End If
	End If
	
	if Int(BallY) < 1 Or Int(BallY) > FrameH - 2 Then
		BallVelocityY = -BallVelocityY
	End If
	
	BallX = BallX + BallVelocityX
	BallY = BallY + BallVelocityY
	If ClearBallTrail And (Int(BallPrevX) <> Int(BallX) Or Int(BallPrevY) <> Int(BallY)) Then
		PongSheet.getCellByPosition(Int(BallPrevX), Int(BallPrevY)).CellBackColor = RGB(255, 240, 240)
		PongSheet.getCellByPosition(Int(BallX), Int(BallY)).CellBackColor = RGB(255, 0, 0)
	End If
		
WEnd
End Sub

Function UpdateBallVelocity(ByRef BallAngle as Double, ByRef BallVelocityX as Double, ByRef BallVelocityY as Double)
	BallVelocityX = Cos(BallAngle)
	BallVelocityY = -Sin(BallAngle)
End Function

Function ResetBall(ByRef BallX as Double, ByRef BallY as Double, ByRef BallAngle as Double, ByRef BallVelocityX, ByRef BallVelocityY)
	BallX = Int(FrameW / 2)
	BallY = Int(FrameH / 2)
	BallAngle = Int((10 * Rnd) + 20)	
	UpdateBallVelocity(BallAngle, BallVelocityX, BallVelocityY)
End Function

Function ClearAndRedraw(ByRef PongSheet as Object, PaddleX as Integer, PaddleY as Integer, PaddleH as Integer, Paddle2X as Integer, Paddle2Y as Integer, Paddle2H as Integer)
	PongSheet.getCellRangeByName("A1:BZ50").CellBackColor = RGB(255, 255, 255)
	DrawPaddle(PongSheet, PaddleX, PaddleY, PaddleH)
	DrawPaddle(PongSheet, Paddle2X, Paddle2Y, Paddle2H)
End Function

Function DrawPaddle(ByRef PongSheet as Object, ByRef PaddleX as Integer, ByRef PaddleY as Integer, ByRef PaddleH as Integer)
	Dim i as Integer
	For i = 0 To PaddleH
		PongSheet.getCellByPosition(PaddleX, PaddleY + i).CellBackColor = RGB(0, 0, 0)
	Next i
End Function

Function KeyHandler_keyPressed(Event as Object) as Boolean
	If PrevKey = Event.KeyCode Then
		Exit Function
	End If
	If Event.KeyCode = com.sun.star.awt.Key.UP Then
		PaddleUp = True
	End If
	If Event.KeyCode = com.sun.star.awt.Key.DOWN Then
		PaddleDown = True
	End If
	PrevKey = Event.KeyCode
	KeyHandler_keyPressed = False
End Function

Function KeyHandler_keyReleased(Event as Object) as Boolean
	If Event.KeyCode = com.sun.star.awt.Key.UP Then
		PaddleUp = False
		PrevKey = Empty
	End If
	If Event.KeyCode = com.sun.star.awt.Key.DOWN Then
		PaddleDown = False
		PrevKey = Empty
	End If
	KeyHandler_keyReleased = False
End Function

