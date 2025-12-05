REM *** libreoffice-pong by SammygoodTunes (2025) ***

Option Explicit

Global KeyHandler
Global PaddleUp
Global PaddleDown
Global PrevKey

Global FrameW
Global FrameH

Global Start
Global ClearScreenFlag
Global DrawTitleFlag
Global DrawObstacleFlag

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

Dim ObstacleX As Integer
Dim ObstacleY As Integer
Dim ObstacleH As Integer
Dim ObstacleDuration As Integer

Dim PlayerScoreCell As Object
Dim OpponentScoreCell As Object
Dim PlayerScore As Integer
Dim OpponentScore As Integer

Dim ClearBallTrail As Boolean
Dim OpponentDelay As Double
Dim Diff As Double

Dim PongSheet As Object, Cell As Object

Start = False
ClearScreenFlag = False
DrawTitleFlag = True
DrawObstacleFlag = False

ObstacleX = 0
ObstacleY = 0
ObstacleH = 0
ObstacleDuration = 0

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

PlayerScoreCell = PongSheet.getCellRangeByName("CA51")
OpponentScoreCell = PongSheet.getCellRangeByName("CB51")
PlayerScoreCell.Columns.Width = 2700
PlayerScoreCell.String = "Player Score: " & PlayerScore
OpponentScoreCell.Columns.Width = 3200
OpponentScoreCell.String = "Opponent Score: " & OpponentScore

MsgBox "Pong by SammygoodTunes (2025) - Press SPACE on title screen to start"

While True
	If ClearScreenFlag Then
		ClearAndRedraw(PongSheet, PaddleX, PaddleY, PaddleH, Paddle2X, Paddle2Y, Paddle2H)
		ClearScreenFlag = False
	End If

	If Start Then
		
		ClearBallTrail = True
		BallPrevX = BallX
		BallPrevY = BallY
		ThisComponent.getCurrentController().select(PongSheet.getCellByPosition(0, 0))
		
		If Int(600 * Rnd) = 0 And ObstacleDuration <= 0 Then
			ObstacleX = Int(((FrameW \ 3) * Rnd) + FrameW \ 3)
			ObstacleH = Int((FrameH - FrameH \ 4 - 10) * Rnd + 10)
			ObstacleY = Int((FrameH - ObstacleH - 1) * Rnd)
			ObstacleDuration = Int(750 * Rnd + 250)
			DrawObstacleFlag = True
		End If
		
	
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
		
		If Paddle2Y > 0 And BallY < Paddle2Y + Paddle2H / 2.5 And BallX > FrameW \ 2 + OpponentDelay Then
			Paddle2Y = Paddle2Y - 1
			PongSheet.getCellByPosition(Paddle2X, Paddle2Y).CellBackColor = RGB(0, 0, 0)
			PongSheet.getCellByPosition(Paddle2X, Paddle2Y + Paddle2H + 1).CellBackColor = RGB(255, 255, 255)
	
		End If
		
		If Paddle2Y < FrameH - 1 - Paddle2H And BallY > Paddle2Y + Paddle2H / 1.5 And BallX > FrameW \ 2 + OpponentDelay Then
			Paddle2Y = Paddle2Y + 1
			PongSheet.getCellByPosition(Paddle2X, Paddle2Y - 1).CellBackColor = RGB(255, 255, 255)
			PongSheet.getCellByPosition(Paddle2X, Paddle2Y + Paddle2H).CellBackColor = RGB(0, 0, 0)
		End If
		
		
		If Int(BallX) < 1 Then
			If Int(BallY) < PaddleY Or Int(BallY) > PaddleY + PaddleH + 1 Then
				PongSheet.getCellByPosition(Int(BallX), Int(BallY)).CellBackColor = RGB(255, 255, 255)
				OpponentScore = OpponentScore + 1
				ResetBall(BallX, BallY, BallAngle, BallVelocityX, BallVelocityY)
				PongSheet.getCellRangeByName("CB51").String = "Opponent Score: " & OpponentScore
				If BallVelocityX < 0 Then
					BallVelocityX = -BallVelocityX
				EndIf
				ClearAndRedraw(PongSheet, PaddleX, PaddleY, PaddleH, Paddle2X, Paddle2Y, Paddle2H)
				If ObstacleDuration > 0 Then
					DrawObstacle(PongSheet, ObstacleX, ObstacleY, ObstacleH)
				End If
			Else
				Diff = ((PaddleY + PaddleH / 2) - BallY) / (PaddleH / 2)
				ClearBallTrail = False
				PongSheet.getCellByPosition(Int(BallX), Int(BallY)).CellBackColor = RGB(0, 0, 0)
				BallAngle = (5 * 3.141592 / 12) * Diff
				UpdateBallVelocity(BallAngle, BallVelocityX, BallVelocityY)
				If BallVelocityX < 0 Then
					BallVelocityX = -BallVelocityX
				EndIf
				OpponentDelay = (FrameW / 3.25 * Rnd)
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
				If ObstacleDuration > 0 Then
					DrawObstacle(PongSheet, ObstacleX, ObstacleY, ObstacleH)
				End If
			Else
				Diff = ((Paddle2Y + Paddle2H / 2) - BallY) / (Paddle2H / 2)
				ClearBallTrail = False
				PongSheet.getCellByPosition(Int(BallX), Int(BallY)).CellBackColor = RGB(0, 0, 0)
				BallAngle = (5 * 3.141592 / 12) * Diff
				UpdateBallVelocity(BallAngle, BallVelocityX, BallVelocityY)
				BallVelocityX = -BallVelocityX
			End If
		End If
		
		If Int(BallY) < 1 Or Int(BallY) > FrameH - 2 Then
			BallVelocityY = -BallVelocityY
		End If
		
		BallX = BallX + BallVelocityX
		BallY = BallY + BallVelocityY
		
		If ClearBallTrail And (Int(BallPrevX) <> Int(BallX) Or Int(BallPrevY) <> Int(BallY)) Then
			PongSheet.getCellByPosition(Int(BallPrevX), Int(BallPrevY)).CellBackColor = RGB(255, 240, 240)
			PongSheet.getCellByPosition(Int(BallX), Int(BallY)).CellBackColor = RGB(255, 0, 0)
		End If
		
		
		If ObstacleDuration > 0 And Int(BallPrevY) >= ObstacleY - 1 And Int(BallPrevY) <= ObstacleY + ObstacleH + 1 And Int(BallPrevX) = ObstacleX Then
			If Int(BallY) <> ObstacleY - 1 And Int(BallY) <> ObstacleY + ObstacleH + 1 Then
				PongSheet.getCellByPosition(Int(BallPrevX), Int(BallPrevY)).CellBackColor = RGB(150, 150, 150)
			Endif
		End If
		
		If ObstacleDuration > 0 And Int(BallY) >= ObstacleY - 1 And Int(BallY) <= ObstacleY + ObstacleH + 1 And Int(BallX) = ObstacleX Then
			If Int(BallY) = ObstacleY - 1 Or Int(BallY) = ObstacleY + ObstacleH + 1 Then
				BallVelocityY = -BallVelocityY
			Else
				BallVelocityX = -BallVelocityX
				PongSheet.getCellByPosition(Int(BallX), Int(BallY)).CellBackColor = RGB(150, 150, 150)
			Endif
		End If
		
		
		
		If ObstacleDuration > 0 Then
			ObstacleDuration = ObstacleDuration - 1
			If ObstacleDuration = 0 Then
				ClearObstacle(PongSheet, ObstacleX, ObstacleY, ObstacleH)
			End If
		End If

		If DrawObstacleFlag Then
			If ObstacleDuration > 0 Then
				DrawObstacle(PongSheet, ObstacleX, ObstacleY, ObstacleH)
			End If
			DrawObstacleFlag = False
		End If
	Else
		If DrawTitleFlag Then
			DrawTitleFlag = False
			DrawTitle(PongSheet)
		End If
	End If
	Wait 25
WEnd
End Sub

Function UpdateBallVelocity(ByRef BallAngle As Double, ByRef BallVelocityX As Double, ByRef BallVelocityY As Double)
	BallVelocityX = Cos(BallAngle)
	BallVelocityY = -Sin(BallAngle)
End Function

Function ResetBall(ByRef BallX As Double, ByRef BallY As Double, ByRef BallAngle As Double, ByRef BallVelocityX, ByRef BallVelocityY)
	BallX = Int(FrameW / 2)
	BallY = Int(FrameH / 2)
	BallAngle = Int((10 * Rnd) + 20)	
	UpdateBallVelocity(BallAngle, BallVelocityX, BallVelocityY)
End Function

Function ClearAndRedraw(ByRef PongSheet As Object, PaddleX As Integer, PaddleY As Integer, PaddleH As Integer, Paddle2X As Integer, Paddle2Y As Integer, Paddle2H As Integer)
	PongSheet.getCellRangeByName("A1:BZ50").CellBackColor = RGB(255, 255, 255)
	DrawPaddle(PongSheet, PaddleX, PaddleY, PaddleH)
	DrawPaddle(PongSheet, Paddle2X, Paddle2Y, Paddle2H)
End Function

Function DrawPaddle(ByRef PongSheet As Object, ByRef PaddleX As Integer, ByRef PaddleY As Integer, ByRef PaddleH As Integer)
	Dim i As Integer
	For i = 0 To PaddleH
		PongSheet.getCellByPosition(PaddleX, PaddleY + i).CellBackColor = RGB(0, 0, 0)
	Next i
End Function

Function DrawObstacle(ByRef PongSheet As Object, ByRef ObstacleX As Integer, ByRef ObstacleY As Integer, ByRef ObstacleH As Integer)
	Dim i As Integer
	For i = 0 To ObstacleH
		PongSheet.getCellByPosition(ObstacleX, ObstacleY + i).CellBackColor = RGB(150, 150, 150)
	Next i
End Function

Function ClearObstacle(ByRef PongSheet As Object, ByRef ObstacleX As Integer, ByRef ObstacleY As Integer, ByRef ObstacleH As Integer)
	Dim i As Integer
	For i = 0 To ObstacleH
		PongSheet.getCellByPosition(ObstacleX, ObstacleY + i).CellBackColor = RGB(255, 255, 255)
	Next i
End Function


Function DrawTitle(ByRef PongSheet As Object)
	PongSheet.getCellRangeByName("A1:BZ50").CellBackColor = RGB(255, 255, 255)
	PongSheet.getCellRangeByName("AC11:AC16").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AD11:AE11").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AD14:AE14").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AF12:AF13").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AH14:AH15").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AI13:AJ13").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AK14:AK15").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AI16:AJ16").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AM13:AM16").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AN14").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AO13").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AP14:AP16").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AR14:AR15").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AS13:AU13").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AU14:AU17").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AS16:AT16").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AS18:AT18").CellBackColor = RGB(0, 0, 0)
		
	PongSheet.getCellRangeByName("G20:G24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("H20").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("I21").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("H22").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("I23").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("H24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("K20:K21").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("M20:M21").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("L22:L24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("S20:T20").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("R21").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("S22").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("T23").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("R24:S24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("V21:V24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("W20").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("X21:X24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("W22").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("Z20:Z24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AA21:AA22").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AB20:AB24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AD20:AD24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AE21:AE22").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AF20:AF24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AH20:AH21").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AJ20:AJ21").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AI22:AI24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AM20:AN20").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AL21:AL23").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AN22:AN24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AM24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AP21:AP23").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AQ20").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AR21:AR23").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AQ24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AT21:AT23").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AU20").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AV21:AV23").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AU24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AX20:AX24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AY20").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AY24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AZ21:AZ23").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BB20:BD20").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BC21:BC24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BF20:BF23").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BG24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BH20:BH23").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BJ20:BJ24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BK21").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BL22").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BM20:BM24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BO20:BO24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BP20").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BP22").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BP24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BS20:BT20").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BR21").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BS22").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BT23").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("BR24:BS24").CellBackColor = RGB(0, 0, 0)
	PongSheet.getCellRangeByName("AC9:AC10").CellBackColor = RGB(255, 240, 240)
	PongSheet.getCellRangeByName("AD8").CellBackColor = RGB(255, 240, 240)
	PongSheet.getCellRangeByName("AE7:AF7").CellBackColor = RGB(255, 240, 240)
	PongSheet.getCellRangeByName("AG8").CellBackColor = RGB(255, 240, 240)
	PongSheet.getCellRangeByName("AH9:AH10").CellBackColor = RGB(255, 240, 240)
	PongSheet.getCellRangeByName("AI11:AI12").CellBackColor = RGB(255, 240, 240)
	PongSheet.getCellRangeByName("AJ9:AJ10").CellBackColor = RGB(255, 240, 240)
	PongSheet.getCellRangeByName("AK8:AL8").CellBackColor = RGB(255, 240, 240)
	PongSheet.getCellRangeByName("AM9:AM10").CellBackColor = RGB(255, 240, 240)
	PongSheet.getCellRangeByName("AN11:AN13").CellBackColor = RGB(255, 240, 240)
	PongSheet.getCellRangeByName("AO10").CellBackColor = RGB(255, 240, 240)
	PongSheet.getCellRangeByName("AP9:AQ9").CellBackColor = RGB(255, 240, 240)
	PongSheet.getCellRangeByName("AR10:AR11").CellBackColor = RGB(255, 240, 240)
	PongSheet.getCellRangeByName("AS12").CellBackColor = RGB(255, 240, 240)
	PongSheet.getCellRangeByName("AT11").CellBackColor = RGB(255, 240, 240)
	PongSheet.getCellRangeByName("AU12").CellBackColor = RGB(255, 0, 0)
	
	PongSheet.getCellRangeByName("G26:BT26").CellBackColor = RGB(50, 50, 50)
	PongSheet.getCellRangeByName("G27:BT27").CellBackColor = RGB(0, 0, 0)
End Function

Function KeyHandler_keyPressed(Event As Object) As Boolean
	If PrevKey = Event.KeyCode Then
		Exit Function
	End If
	If Event.KeyCode = com.sun.star.awt.Key.UP Then
		PaddleUp = True
	End If
	If Event.KeyCode = com.sun.star.awt.Key.DOWN Then
		PaddleDown = True
	End If
	If Event.KeyCode = com.sun.star.awt.Key.SPACE Then
		Start = Not Start
		ClearScreenFlag = Start
		DrawTitleFlag = Not Start
		DrawObstacleFlag = Start
	End If
	PrevKey = Event.KeyCode
	KeyHandler_keyPressed = False
End Function

Function KeyHandler_keyReleased(Event As Object) As Boolean
	If Event.KeyCode = com.sun.star.awt.Key.UP Then
		PaddleUp = False
		PrevKey = Empty
	End If
	If Event.KeyCode = com.sun.star.awt.Key.DOWN Then
		PaddleDown = False
		PrevKey = Empty
	End If
	If Event.KeyCode = com.sun.star.awt.Key.SPACE Then
		PrevKey = Empty
	End If
	KeyHandler_keyReleased = False
End Function
