
Option Explicit

Global keyHandler
Global paddleUp
Global paddleDown
Global PrevKey

Sub Main

Const MAX_ANGLE = 2 * 3.14159265358979 / 12

Dim paddleX As Integer
Dim paddleY As Integer
Dim paddleW As Integer
Dim paddleH As Integer

Dim paddle2X As Integer
Dim paddle2Y As Integer
Dim paddle2W As Integer
Dim paddle2H As Integer

Dim ballX As Double
Dim ballY As Double
Dim ballVelocityX As Double
Dim ballVelocityY As Double
Dim ballAngle As Double

Dim playerScore as Integer
Dim opponentScore as Integer

Dim clearBallTrail as Boolean

playerScore = 0
opponentScore = 0

paddleX = 0
paddleY = 0
paddleH = 6

paddle2X = 77
paddle2Y = 0
paddle2H = 6

ballX = 77 \ 2
ballY = 50 \ 2
ballAngle = 45

ballVelocityX = Cos(ballAngle)
ballVelocityY = -Sin(ballAngle)

If ballVelocityX = 0 Then 
	ballVelocityX = -1
End If
If ballVelocityY = 0 Then 
	ballVelocityY = -1
End If

paddleUp = False
paddleDown = False

Dim PongSheet As Object, Cell As Object

keyHandler = CreateUnoListener("KeyHandler_", "com.sun.star.awt.XKeyHandler")
ThisComponent.CurrentController.addKeyHandler(keyHandler)

PongSheet = ThisComponent.Sheets.getByName("PONG")
PongSheet.getCellRangeByName("A1:BZ50").CellBackColor = RGB(255, 255, 255)
PongSheet.getCellRangeByName("A1:BZ50").Rows.Height = 200
PongSheet.getCellRangeByName("A1:BZ50").Columns.Width = 200

DrawPaddle(PongSheet, paddleX, paddleY, paddleH)
DrawPaddle(PongSheet, paddle2X, paddle2Y, paddle2H)

PongSheet.getCellRangeByName("CA51").String = "Player Score: " & playerScore
PongSheet.getCellRangeByName("CB51").String = "Opponent Score: " & opponentScore

While True
	clearBallTrail = True
	ThisComponent.getCurrentController().select(PongSheet.getCellByPosition(0, 0))

	If paddleUp And paddleY > 0 Then
		paddleY = paddleY - 1
		PongSheet.getCellByPosition(paddleX, paddleY).CellBackColor = RGB(0, 0, 0)
		PongSheet.getCellByPosition(paddleX, paddleY + paddleH + 1).CellBackColor = RGB(255, 255, 255)
	End If
	
	If paddleDown And paddleY < 49 - paddleH Then
		paddleY = paddleY + 1
		PongSheet.getCellByPosition(paddleX, paddleY - 1).CellBackColor = RGB(255, 255, 255)
		PongSheet.getCellByPosition(paddleX, paddleY + paddleH).CellBackColor = RGB(0, 0, 0)
	End If
	
	If ballY < paddle2Y And paddle2Y > 0 Then
		paddle2Y = paddle2Y - 1
		PongSheet.getCellByPosition(paddle2X, paddle2Y).CellBackColor = RGB(0, 0, 0)
		PongSheet.getCellByPosition(paddle2X, paddle2Y + paddle2H + 1).CellBackColor = RGB(255, 255, 255)

	End If
	
	If ballY > paddle2Y + paddle2H And paddle2Y < 49 - paddle2H Then
		paddle2Y = paddle2Y + 1
		PongSheet.getCellByPosition(paddle2X, paddle2Y - 1).CellBackColor = RGB(255, 255, 255)
		PongSheet.getCellByPosition(paddle2X, paddle2Y + paddle2H).CellBackColor = RGB(0, 0, 0)
	End If
	
	
	If Int(ballX) < 1 Then
		If Int(ballY) < paddleY Or Int(ballY) > paddleY + paddleH Then
			MsgBox "You lose"
			opponentScore = opponentScore + 1
			ballX = 77 \ 2
			ballY = 50 \ 2
			ballAngle = Int((10 * Rnd) + 45)
			ballVelocityX = Cos(ballAngle)
			ballVelocityY = -Sin(ballAngle)
			PongSheet.getCellRangeByName("CB51").String = "Opponent Score: " & opponentScore
		Else
			clearBallTrail = False
			PongSheet.getCellByPosition(Int(ballX), Int(ballY)).CellBackColor = RGB(0, 0, 0)
			ballVelocityX = -ballVelocityX
		End If
	End If
	
	If Int(ballX) > 76 Then
		If Int(ballY) < paddle2Y Or Int(ballY) > paddle2Y + paddle2H Then
			MsgBox "You win"
			playerScore = playerScore + 1
			ballX = 77 \ 2
			ballY = 50 \ 2
			ballAngle = Int((10 * Rnd) + 45)
			ballVelocityX = Cos(ballAngle)
			ballVelocityY = -Sin(ballAngle)
			PongSheet.getCellRangeByName("CA51").String = "Player Score: " & playerScore
		Else
			clearBallTrail = False
			PongSheet.getCellByPosition(Int(ballX), Int(ballY)).CellBackColor = RGB(0, 0, 0)
			ballVelocityX = -ballVelocityX
		End If
	End If
	
	if Int(ballY) < 1 Or Int(ballY) > 48 Then
		ballVelocityY = -ballVelocityY
	End If
	
	ballX = ballX + ballVelocityX
	ballY = ballY + ballVelocityY
	If clearBallTrail Then
		PongSheet.getCellByPosition(Int(ballX - ballVelocityX), Int(ballY - ballVelocityY)).CellBackColor = RGB(255, 255, 255)
	End If
	PongSheet.getCellByPosition(Int(ballX), Int(ballY)).CellBackColor = RGB(0, 0, 0)
		
WEnd
End Sub

Function DrawPaddle(ByRef PongSheet as Object, ByRef paddleX as Integer, ByRef paddleY as Integer, ByRef paddleH as Integer)
	Dim i as Integer
	For i = 0 To paddleH
		PongSheet.getCellByPosition(paddleX, paddleY + i).CellBackColor = RGB(0, 0, 0)
	Next i
End Function

Function KeyHandler_keyPressed(Event as Object) as Boolean
	If PrevKey = Event.KeyCode Then
		Exit Function
	End If
	If Event.KeyCode = com.sun.star.awt.Key.UP Then
		paddleUp = True
	End If
	If Event.KeyCode = com.sun.star.awt.Key.DOWN Then
		paddleDown = True
	End If
	PrevKey = Event.KeyCode
	KeyHandler_keyPressed = False
End Function

Function KeyHandler_keyReleased(Event as Object) as Boolean
	If Event.KeyCode = com.sun.star.awt.Key.UP Then
		paddleUp = False
		PrevKey = Empty
	End If
	If Event.KeyCode = com.sun.star.awt.Key.DOWN Then
		paddleDown = False
		PrevKey = Empty
	End If
	KeyHandler_keyReleased = False
End Function


