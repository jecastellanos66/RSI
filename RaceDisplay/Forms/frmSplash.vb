Imports System.Threading

Public NotInheritable Class frmSplash
    Public Shared _frmSplash As frmSplash = Nothing
    Public Shared myThread As Thread = Nothing

    Private Shared Sub ShowMe()
        _frmSplash = New frmSplash
        Application.Run(_frmSplash)
        '
    End Sub

    Public Shared Sub ShowSplash()
        If (_frmSplash IsNot Nothing) Then Return
        '
        myThread = New Thread(New ThreadStart(AddressOf ShowMe))
        myThread.IsBackground = True
        myThread.SetApartmentState(ApartmentState.STA)
        myThread.Start()
        '
    End Sub

    Public Shared Sub CloseForm()
        myThread.Abort()
        myThread = Nothing
        _frmSplash = Nothing
    End Sub

    Private Sub frmSplash_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Label1.Parent = Me.PictureBox1
        Me.Label2.Parent = Me.PictureBox1
        Me.Label1.BackColor = Color.Transparent
        Me.Label2.BackColor = Color.Transparent
        Me.Refresh()
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class