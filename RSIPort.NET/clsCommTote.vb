Option Strict Off
Option Explicit On
Friend Class clsCommTote
	
	Public ToteBuffer As String
	Private m_intArchFile As Short

    'Nov-08, 2001 EMJ - used for multicasting
    Public WithEvents CommTote As MSCommLib.MSComm

    'Private Sub CommTote_OnComm(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CommTote.OnComm
    Private Sub CommTote_OnComm() Handles CommTote.OnComm
        On Error GoTo ErrHndlr

        'part for comm tote
        Dim temp As String
        'UPGRADE_WARNING: Couldn't resolve default property of object CommTote.Input. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        temp = CommTote.Input
        ToteBuffer = ToteBuffer & temp
        Print(ArchFile, temp)

        Exit Sub

ErrHndlr:
        CommTote.InBufferCount = 0
    End Sub


    Public Property ArchFile() As Short
		Get
			ArchFile = m_intArchFile
		End Get
		Set(ByVal Value As Short)
			m_intArchFile = Value
		End Set
	End Property
End Class