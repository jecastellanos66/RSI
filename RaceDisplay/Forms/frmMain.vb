Public Class frmMain

#Region " Declarations "

    Private myRaceDetails As RaceDetails
    'Private myCommSvr As 

#End Region

#Region " Constructor "

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        '
    End Sub

#End Region

#Region " Form Events "

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        frmSplash.ShowSplash()
        For i As Integer = 0 To 5500000
            Application.DoEvents()
        Next
        '
        frmSplash.CloseForm()
        '
    End Sub

    'Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    '    Application.DoEvents()
    '    Me.lblTime.Text = Format(Now, "hh:mm:ss tt")
    'End Sub

#End Region

#Region " Toolbar Events "

    Private Sub CloseToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub ClearOddsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearOddsToolStripMenuItem1.Click
        For Each ctl As Control In Me.gbOdds.Controls
            If TypeOf ctl Is Windows.Forms.TextBox Then ctl.Text = ""
        Next
        ''
    End Sub

    Private Sub ClearRunningOderToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearRunningOderToolStripMenuItem.Click
        For Each ctl As Control In Me.gbOrder.Controls
            If TypeOf ctl Is Windows.Forms.TextBox Then ctl.Text = ""
        Next
        ''
    End Sub

    Private Sub ClearWinToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearWinToolStripMenuItem.Click
        For Each ctl As Control In Me.gbWin.Controls
            If TypeOf ctl Is Windows.Forms.TextBox Then ctl.Text = ""
        Next
        '
    End Sub

    Private Sub ClearPlaceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearPlaceToolStripMenuItem.Click
        For Each ctl As Control In Me.gbPlace.Controls
            If TypeOf ctl Is Windows.Forms.TextBox Then ctl.Text = ""
        Next
        '
    End Sub

    Private Sub ClearShowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearShowToolStripMenuItem.Click
        For Each ctl As Control In Me.gbShow.Controls
            If TypeOf ctl Is Windows.Forms.TextBox Then ctl.Text = ""
        Next
        '
    End Sub

#End Region

   
    
End Class