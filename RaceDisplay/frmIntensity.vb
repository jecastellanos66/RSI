Public Class frmIntensity

    Private Sub frmIntensity_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.gbIin.BackColor = Color.Transparent
        Me.gbOut.BackColor = Color.Transparent
        '
        Me.gbIin.Enabled = My.Settings.HasInsideBoard
        '
        For Each ctl As Control In Me.gbIin.Controls
            If (TypeOf ctl Is Windows.Forms.RadioButton) AndAlso _
               ctl.Tag = My.Settings.InsideIntensity Then
                CType(ctl, Windows.Forms.RadioButton).Checked = True
                Exit For
            End If
        Next
        '
        For Each ctl As Control In Me.gbOut.Controls
            If (TypeOf ctl Is Windows.Forms.RadioButton) AndAlso _
               ctl.Tag = My.Settings.OutsideIntensity Then
                CType(ctl, Windows.Forms.RadioButton).Checked = True
                Exit For
            End If
        Next
        '
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        For Each ctl As Control In Me.gbIin.Controls
            If (TypeOf ctl Is Windows.Forms.RadioButton) AndAlso _
               (CType(ctl, Windows.Forms.RadioButton).Checked) Then
                My.Settings.InsideIntensity = ctl.Tag
            End If
        Next
        '
        For Each ctl As Control In Me.gbOut.Controls
            If (TypeOf ctl Is Windows.Forms.RadioButton) AndAlso _
               (CType(ctl, Windows.Forms.RadioButton).Checked) Then
                My.Settings.OutsideIntensity = ctl.Tag
            End If
        Next
        '
        Me.Close()
    End Sub
End Class