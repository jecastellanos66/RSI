Public Class frmAbout

    Private Sub frmAbout_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Label1.BackColor = Color.Transparent
        Me.Label2.BackColor = Color.Transparent
        Me.Label3.BackColor = Color.Transparent
        Me.LinkLabel1.BackColor = Color.Transparent
        '
        Me.Label2.Text = "Warning: This computer program is protected by copyright law and" & vbCrLf & _
                         "international treaties. Unauthorized reproduction or distribution of" & vbCrLf & _
                         "this program, or any portion of it, may result in severe civil and" & vbCrLf & _
                         "criminal penalties, and will be prosecuted under the maximum extent" & vbCrLf & _
                         "possible under law."
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        'System.Diagnostics.Process.Start("http://www.televiewracing.com/")
    End Sub
End Class