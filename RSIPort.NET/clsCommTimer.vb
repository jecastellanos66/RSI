Option Strict Off
Option Explicit On
Friend Class clsCommTimer
	
	Public TimerBuffer As String
	Private m_intArchTimer As Short
	
	'Nov-08, 2001 EMJ - used for multicasting
	Public WithEvents CommTimer As MSCommLib.MSComm

    'Private Sub CommTimer_OnComm(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CommTimer.OnComm
    Private Sub CommTimer_OnComm() Handles CommTimer.OnComm
        Dim temp As String
        Dim ZeroPsn As Short
        Dim TempLen As Short

        'section for comm timer
        'UPGRADE_WARNING: Couldn't resolve default property of object CommTimer.Input. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        temp = CommTimer.Input
        '    Do Until InStr(temp, Chr$(0)) = 0
        '      ZeroPsn = InStr(temp, Chr$(0))
        '      TempLen = Len(temp)
        '      temp = Left$(temp, ZeroPsn - 1) & " " & Right$(temp, TempLen - ZeroPsn)
        '    Loop
        TimerBuffer = TimerBuffer & temp

        'Print #ArchTimer, temp;

    End Sub


    Public Property ArchTimer() As Short
		Get
			ArchTimer = m_intArchTimer
		End Get
		Set(ByVal Value As Short)
			m_intArchTimer = Value
		End Set
	End Property
End Class