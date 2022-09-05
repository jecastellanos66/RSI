Option Strict Off
Option Explicit On
Friend Class clsTimer2
	
	Public WithEvents Timer2 As ccrpTimers6.ccrpTimer
	
	'Nov-08, 2001 EMJ - used for multicasting
	Public Event TimerTriggered(ByRef fTriggered As Boolean)
	
	Private Sub Timer2_Timer(ByVal Milliseconds As Integer) Handles Timer2.Timer
		
		RaiseEvent TimerTriggered(True)
		
	End Sub
End Class