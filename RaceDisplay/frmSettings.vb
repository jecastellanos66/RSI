Public Class frmSettings

#Region " Events "
    Private Sub frmSettings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '
        Me.ErrorProvider1.SetError(Me.btnOK, "")
        Me.ErrorProvider1.SetError(Me.txtMiniAddress, "")
        '
        Me.rbOutsideH.Checked = My.Settings.OutsideIntensity = Me.rbOutsideH.Tag
        Me.rbOutsideM.Checked = My.Settings.OutsideIntensity = Me.rbOutsideM.Tag
        Me.rbOutsideL.Checked = My.Settings.OutsideIntensity = Me.rbOutsideL.Tag
        Me.rbOutsideMH.Checked = My.Settings.OutsideIntensity = Me.rbOutsideMH.Tag
        Me.rbOutsideML.Checked = My.Settings.OutsideIntensity = Me.rbOutsideML.Tag

        '
        Me.rbInsideYes.Checked = My.Settings.HasInsideBoard
        Me.rbInsideNo.Checked = Not My.Settings.HasInsideBoard
        '
        Me.gbIin.Visible = Me.rbInsideYes.Checked
        '
        If Me.gbIin.Visible Then
            Me.rbInsideH.Checked = My.Settings.InsideIntensity = Me.rbInsideH.Tag
            Me.rbInsideM.Checked = My.Settings.InsideIntensity = Me.rbInsideM.Tag
            Me.rbInsideL.Checked = My.Settings.InsideIntensity = Me.rbInsideL.Tag
        End If
        '
        Me.rbMiniYes.Checked = My.Settings.HasMiniBoard
        Me.rbMiniNo.Checked = Not My.Settings.HasMiniBoard
        '
        'Me.gbMiniAddress.Visible = Me.rbMiniYes.Checked
        Me.gbMiniBrightness.Visible = Me.rbMiniYes.Checked
        '
        If Not Me.rbMiniYes.Checked Then Return
        '
        Me.txtMiniAddress.Text = My.Settings.MiniBoardAddress
        Me.txtMiniBrightness.Value = My.Settings.MiniBoardDimming
        ''
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
        ''
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.ErrorProvider1.SetError(Me.btnOK, "")
        Me.ErrorProvider1.SetError(Me.txtMiniAddress, "")
        '
        If Not Me.AllValuesProvided Then
            Me.ErrorProvider1.SetError(Me.btnOK, "All values have to be provided to continue.")
            Return
        End If
        Me.UpdateSettings()
        '
        Me.Close()
        ''
    End Sub

    Private Sub rbInsideNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbInsideNo.CheckedChanged
        Me.gbIin.Visible = Not Me.rbInsideNo.Checked
    End Sub

    Private Sub rbMiniYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMiniYes.CheckedChanged
        'Me.gbMiniAddress.Visible = Me.rbMiniYes.Checked
        Me.gbMiniBrightness.Visible = Me.rbMiniYes.Checked
    End Sub

    Private Sub rbInsideYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbInsideYes.CheckedChanged
        Me.gbIin.Visible = Me.rbInsideYes.Checked
    End Sub

    Private Sub rbMiniNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMiniNo.CheckedChanged
        'Me.gbMiniAddress.Visible = Me.rbMiniYes.Checked
        Me.gbMiniBrightness.Visible = Me.rbMiniYes.Checked
    End Sub

#End Region

#Region " Methods "

    Private Function AllValuesProvided() As Boolean
        If Not Me.rbOutsideH.Checked AndAlso _
           Not Me.rbOutsideM.Checked AndAlso _
           Not Me.rbOutsideL.Checked AndAlso _
           Not Me.rbOutsideMH.Checked AndAlso _
           Not Me.rbOutsideML.Checked Then Return False
        '
        If Not Me.rbInsideYes.Checked AndAlso _
           Not Me.rbInsideNo.Checked Then Return False
        '
        If Me.rbInsideYes.Checked Then
            If Not Me.rbInsideH.Checked AndAlso _
               Not Me.rbInsideM.Checked AndAlso _
               Not Me.rbInsideL.Checked Then Return False
        End If
        '
        If Not Me.rbMiniYes.Checked AndAlso _
           Not Me.rbMiniNo.Checked Then Return False
        '
        'If Me.rbMiniYes.Checked Then
        '    If (Me.txtMiniAddress.Text.Trim = "") OrElse _
        '       (Not IsNumeric(Me.txtMiniAddress.Text)) OrElse _
        '       (CType(Me.txtMiniAddress.Text, Integer) > 64) OrElse _
        '       (CType(Me.txtMiniAddress.Text, Integer) < 1) Then
        '        '
        '        Me.ErrorProvider1.SetError(Me.txtMiniAddress, "Value should be between 1 and 64.")
        '        Return False
        '    End If
        '    'the mini brightness will always have the correct value, no need to validate
        'End If
        '
        Return True
    End Function

    Private Sub UpdateSettings()
        '
        For Each ctl As Control In Me.gbOut.Controls
            If (TypeOf ctl Is Windows.Forms.RadioButton) AndAlso _
               (CType(ctl, Windows.Forms.RadioButton).Checked) Then
                My.Settings.OutsideIntensity = ctl.Tag
            End If
        Next
        '
        My.Settings.HasInsideBoard = False  'If(Me.rbInsideYes.Checked, True, False)
        '
        If Me.rbInsideYes.Checked Then
            For Each ctl As Control In Me.gbIin.Controls
                If (TypeOf ctl Is Windows.Forms.RadioButton) AndAlso _
                   (CType(ctl, Windows.Forms.RadioButton).Checked) Then
                    My.Settings.InsideIntensity = ctl.Tag
                End If
            Next
        End If
        '
        My.Settings.HasMiniBoard = False 'If(Me.rbMiniYes.Checked, True, False)
        '
        If Not Me.rbMiniYes.Checked Then Return
        '
        My.Settings.MiniBoardAddress = Me.txtMiniAddress.Text
        My.Settings.MiniBoardDimming = Me.txtMiniBrightness.Value
        '
    End Sub

#End Region

End Class