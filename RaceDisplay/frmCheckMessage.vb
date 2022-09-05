Public Class frmCheckMessage
    Dim myCommSvr As CommSvr
    Private myCurrentMessage As String

    Public Sub New(ByVal CommSvr As CommSvr)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        '
        Me.myCommSvr = CommSvr
        AddHandler Me.myCommSvr.DisplayMessages, AddressOf myCommSvr_DisplayMessages

    End Sub

    'Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    '    '
    '    Me.myCurrentMessage = Me.myCommSvr.CurrentMessage
    '    Me.rtbMsg.Text = Me.myCurrentMessage
    '    '
    '    Me.lblMsg.Text = "Checking Data Transfer"
    '    Me.lblMsg.Visible = True
    '    '
    'End Sub


    Private Delegate Sub ResetMessages()

    Private Sub myCommSvr_DisplayMessages()
        If Me.InvokeRequired Then
            Me.Invoke(New ResetMessages(AddressOf myCommSvr_DisplayMessages))
        Else
            '
            Me.myCurrentMessage = Me.myCommSvr.CurrentMessage
            Me.rtbMsg.Text = Me.myCurrentMessage
            '
            Me.lblMsg.Text = "Checking Data Transfer"
            Me.lblMsg.Visible = True
            '
        End If
    End Sub

    Private Sub frmCheckMessage_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.lblMsg.BackColor = Color.Transparent
        Me.lblMsg.Visible = False
    End Sub

End Class