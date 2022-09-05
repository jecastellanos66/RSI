Option Strict Off
Option Explicit On
Friend Class clsTimer
	
	Public WithEvents Timer1 As ccrpTimers6.ccrpTimer
	
	'Nov-08, 2001 EMJ - used for multicasting
	Public Event TimerTriggered(ByRef fTriggered As Boolean)
	
	Private Sub Timer1_Timer(ByVal Milliseconds As Integer) Handles Timer1.Timer
		
		RaiseEvent TimerTriggered(True)
		
	End Sub
End Class