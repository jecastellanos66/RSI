<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSettings
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSettings))
        Me.gbIin = New System.Windows.Forms.GroupBox
        Me.rbInsideL = New System.Windows.Forms.RadioButton
        Me.rbInsideM = New System.Windows.Forms.RadioButton
        Me.rbInsideH = New System.Windows.Forms.RadioButton
        Me.gbOut = New System.Windows.Forms.GroupBox
        Me.rbOutsideMH = New System.Windows.Forms.RadioButton
        Me.rbOutsideML = New System.Windows.Forms.RadioButton
        Me.rbOutsideL = New System.Windows.Forms.RadioButton
        Me.rbOutsideM = New System.Windows.Forms.RadioButton
        Me.rbOutsideH = New System.Windows.Forms.RadioButton
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.rbInsideNo = New System.Windows.Forms.RadioButton
        Me.rbInsideYes = New System.Windows.Forms.RadioButton
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.rbMiniNo = New System.Windows.Forms.RadioButton
        Me.rbMiniYes = New System.Windows.Forms.RadioButton
        Me.gbMiniAddress = New System.Windows.Forms.GroupBox
        Me.txtMiniAddress = New System.Windows.Forms.TextBox
        Me.gbMiniBrightness = New System.Windows.Forms.GroupBox
        Me.txtMiniBrightness = New System.Windows.Forms.NumericUpDown
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.gbIin.SuspendLayout()
        Me.gbOut.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.gbMiniAddress.SuspendLayout()
        Me.gbMiniBrightness.SuspendLayout()
        CType(Me.txtMiniBrightness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gbIin
        '
        Me.gbIin.Controls.Add(Me.rbInsideL)
        Me.gbIin.Controls.Add(Me.rbInsideM)
        Me.gbIin.Controls.Add(Me.rbInsideH)
        Me.gbIin.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbIin.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.gbIin.Location = New System.Drawing.Point(18, 152)
        Me.gbIin.Name = "gbIin"
        Me.gbIin.Size = New System.Drawing.Size(177, 46)
        Me.gbIin.TabIndex = 11
        Me.gbIin.TabStop = False
        Me.gbIin.Text = "Inside Board Intensity"
        Me.gbIin.Visible = False
        '
        'rbInsideL
        '
        Me.rbInsideL.AutoSize = True
        Me.rbInsideL.Location = New System.Drawing.Point(127, 19)
        Me.rbInsideL.Name = "rbInsideL"
        Me.rbInsideL.Size = New System.Drawing.Size(53, 20)
        Me.rbInsideL.TabIndex = 9
        Me.rbInsideL.Tag = "3F"
        Me.rbInsideL.Text = "Low"
        Me.rbInsideL.UseVisualStyleBackColor = True
        '
        'rbInsideM
        '
        Me.rbInsideM.AutoSize = True
        Me.rbInsideM.Location = New System.Drawing.Point(59, 19)
        Me.rbInsideM.Name = "rbInsideM"
        Me.rbInsideM.Size = New System.Drawing.Size(80, 20)
        Me.rbInsideM.TabIndex = 8
        Me.rbInsideM.Tag = "26"
        Me.rbInsideM.Text = "Medium"
        Me.rbInsideM.UseVisualStyleBackColor = True
        '
        'rbInsideH
        '
        Me.rbInsideH.AutoSize = True
        Me.rbInsideH.Location = New System.Drawing.Point(6, 19)
        Me.rbInsideH.Name = "rbInsideH"
        Me.rbInsideH.Size = New System.Drawing.Size(58, 20)
        Me.rbInsideH.TabIndex = 7
        Me.rbInsideH.Tag = "22"
        Me.rbInsideH.Text = "High"
        Me.rbInsideH.UseVisualStyleBackColor = True
        '
        'gbOut
        '
        Me.gbOut.Controls.Add(Me.rbOutsideMH)
        Me.gbOut.Controls.Add(Me.rbOutsideML)
        Me.gbOut.Controls.Add(Me.rbOutsideL)
        Me.gbOut.Controls.Add(Me.rbOutsideM)
        Me.gbOut.Controls.Add(Me.rbOutsideH)
        Me.gbOut.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbOut.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.gbOut.Location = New System.Drawing.Point(12, 25)
        Me.gbOut.Name = "gbOut"
        Me.gbOut.Size = New System.Drawing.Size(247, 90)
        Me.gbOut.TabIndex = 12
        Me.gbOut.TabStop = False
        Me.gbOut.Text = "Board Brightness"
        '
        'rbOutsideMH
        '
        Me.rbOutsideMH.AutoSize = True
        Me.rbOutsideMH.Location = New System.Drawing.Point(6, 45)
        Me.rbOutsideMH.Name = "rbOutsideMH"
        Me.rbOutsideMH.Size = New System.Drawing.Size(116, 20)
        Me.rbOutsideMH.TabIndex = 6
        Me.rbOutsideMH.TabStop = True
        Me.rbOutsideMH.Tag = "26"
        Me.rbOutsideMH.Text = "Medium High"
        Me.rbOutsideMH.UseVisualStyleBackColor = True
        '
        'rbOutsideML
        '
        Me.rbOutsideML.AutoSize = True
        Me.rbOutsideML.Location = New System.Drawing.Point(130, 47)
        Me.rbOutsideML.Name = "rbOutsideML"
        Me.rbOutsideML.Size = New System.Drawing.Size(111, 20)
        Me.rbOutsideML.TabIndex = 5
        Me.rbOutsideML.TabStop = True
        Me.rbOutsideML.Tag = "36"
        Me.rbOutsideML.Text = "Medium Low"
        Me.rbOutsideML.UseVisualStyleBackColor = True
        '
        'rbOutsideL
        '
        Me.rbOutsideL.AutoSize = True
        Me.rbOutsideL.Location = New System.Drawing.Point(156, 21)
        Me.rbOutsideL.Name = "rbOutsideL"
        Me.rbOutsideL.Size = New System.Drawing.Size(53, 20)
        Me.rbOutsideL.TabIndex = 4
        Me.rbOutsideL.TabStop = True
        Me.rbOutsideL.Tag = "3F"
        Me.rbOutsideL.Text = "Low"
        Me.rbOutsideL.UseVisualStyleBackColor = True
        '
        'rbOutsideM
        '
        Me.rbOutsideM.AutoSize = True
        Me.rbOutsideM.Location = New System.Drawing.Point(70, 19)
        Me.rbOutsideM.Name = "rbOutsideM"
        Me.rbOutsideM.Size = New System.Drawing.Size(80, 20)
        Me.rbOutsideM.TabIndex = 3
        Me.rbOutsideM.TabStop = True
        Me.rbOutsideM.Tag = "2E"
        Me.rbOutsideM.Text = "Medium"
        Me.rbOutsideM.UseVisualStyleBackColor = True
        '
        'rbOutsideH
        '
        Me.rbOutsideH.AutoSize = True
        Me.rbOutsideH.Location = New System.Drawing.Point(6, 19)
        Me.rbOutsideH.Name = "rbOutsideH"
        Me.rbOutsideH.Size = New System.Drawing.Size(58, 20)
        Me.rbOutsideH.TabIndex = 2
        Me.rbOutsideH.TabStop = True
        Me.rbOutsideH.Tag = "22"
        Me.rbOutsideH.Text = "High"
        Me.rbOutsideH.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rbInsideNo)
        Me.GroupBox2.Controls.Add(Me.rbInsideYes)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.GroupBox2.Location = New System.Drawing.Point(15, 160)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(98, 46)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Inside Board"
        '
        'rbInsideNo
        '
        Me.rbInsideNo.AutoSize = True
        Me.rbInsideNo.Location = New System.Drawing.Point(52, 19)
        Me.rbInsideNo.Name = "rbInsideNo"
        Me.rbInsideNo.Size = New System.Drawing.Size(46, 20)
        Me.rbInsideNo.TabIndex = 6
        Me.rbInsideNo.TabStop = True
        Me.rbInsideNo.Text = "No"
        Me.rbInsideNo.UseVisualStyleBackColor = True
        '
        'rbInsideYes
        '
        Me.rbInsideYes.AutoSize = True
        Me.rbInsideYes.Location = New System.Drawing.Point(3, 19)
        Me.rbInsideYes.Name = "rbInsideYes"
        Me.rbInsideYes.Size = New System.Drawing.Size(53, 20)
        Me.rbInsideYes.TabIndex = 5
        Me.rbInsideYes.TabStop = True
        Me.rbInsideYes.Text = "Yes"
        Me.rbInsideYes.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.rbMiniNo)
        Me.GroupBox3.Controls.Add(Me.rbMiniYes)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.GroupBox3.Location = New System.Drawing.Point(134, 160)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(98, 52)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Mini Board"
        Me.GroupBox3.Visible = False
        '
        'rbMiniNo
        '
        Me.rbMiniNo.AutoSize = True
        Me.rbMiniNo.Location = New System.Drawing.Point(52, 20)
        Me.rbMiniNo.Name = "rbMiniNo"
        Me.rbMiniNo.Size = New System.Drawing.Size(46, 20)
        Me.rbMiniNo.TabIndex = 11
        Me.rbMiniNo.TabStop = True
        Me.rbMiniNo.Text = "No"
        Me.rbMiniNo.UseVisualStyleBackColor = True
        '
        'rbMiniYes
        '
        Me.rbMiniYes.AutoSize = True
        Me.rbMiniYes.Location = New System.Drawing.Point(3, 20)
        Me.rbMiniYes.Name = "rbMiniYes"
        Me.rbMiniYes.Size = New System.Drawing.Size(53, 20)
        Me.rbMiniYes.TabIndex = 10
        Me.rbMiniYes.TabStop = True
        Me.rbMiniYes.Text = "Yes"
        Me.rbMiniYes.UseVisualStyleBackColor = True
        '
        'gbMiniAddress
        '
        Me.gbMiniAddress.Controls.Add(Me.txtMiniAddress)
        Me.gbMiniAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbMiniAddress.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.gbMiniAddress.Location = New System.Drawing.Point(279, 164)
        Me.gbMiniAddress.Name = "gbMiniAddress"
        Me.gbMiniAddress.Size = New System.Drawing.Size(114, 52)
        Me.gbMiniAddress.TabIndex = 3
        Me.gbMiniAddress.TabStop = False
        Me.gbMiniAddress.Text = "Mini Board Address"
        Me.gbMiniAddress.Visible = False
        '
        'txtMiniAddress
        '
        Me.ErrorProvider1.SetError(Me.txtMiniAddress, "1")
        Me.txtMiniAddress.Location = New System.Drawing.Point(6, 20)
        Me.txtMiniAddress.MaxLength = 2
        Me.txtMiniAddress.Name = "txtMiniAddress"
        Me.txtMiniAddress.Size = New System.Drawing.Size(54, 22)
        Me.txtMiniAddress.TabIndex = 12
        Me.txtMiniAddress.Text = "10"
        '
        'gbMiniBrightness
        '
        Me.gbMiniBrightness.Controls.Add(Me.txtMiniBrightness)
        Me.gbMiniBrightness.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbMiniBrightness.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.gbMiniBrightness.Location = New System.Drawing.Point(128, 164)
        Me.gbMiniBrightness.Name = "gbMiniBrightness"
        Me.gbMiniBrightness.Size = New System.Drawing.Size(129, 52)
        Me.gbMiniBrightness.TabIndex = 4
        Me.gbMiniBrightness.TabStop = False
        Me.gbMiniBrightness.Text = "Mini Board Brightness"
        Me.gbMiniBrightness.Visible = False
        '
        'txtMiniBrightness
        '
        Me.txtMiniBrightness.Location = New System.Drawing.Point(6, 20)
        Me.txtMiniBrightness.Maximum = New Decimal(New Integer() {15, 0, 0, 0})
        Me.txtMiniBrightness.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.txtMiniBrightness.Name = "txtMiniBrightness"
        Me.txtMiniBrightness.Size = New System.Drawing.Size(51, 22)
        Me.txtMiniBrightness.TabIndex = 13
        Me.txtMiniBrightness.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'btnOK
        '
        Me.ErrorProvider1.SetError(Me.btnOK, "1")
        Me.btnOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.ForeColor = System.Drawing.Color.Black
        Me.ErrorProvider1.SetIconAlignment(Me.btnOK, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.btnOK.Location = New System.Drawing.Point(279, 41)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(54, 23)
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.Color.Black
        Me.btnCancel.Location = New System.Drawing.Point(339, 41)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(71, 23)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "&Close"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'frmSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSlateGray
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(422, 130)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.gbMiniAddress)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.gbIin)
        Me.Controls.Add(Me.gbMiniBrightness)
        Me.Controls.Add(Me.gbOut)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSettings"
        Me.Text = "Settings"
        Me.gbIin.ResumeLayout(False)
        Me.gbIin.PerformLayout()
        Me.gbOut.ResumeLayout(False)
        Me.gbOut.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.gbMiniAddress.ResumeLayout(False)
        Me.gbMiniAddress.PerformLayout()
        Me.gbMiniBrightness.ResumeLayout(False)
        CType(Me.txtMiniBrightness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbIin As System.Windows.Forms.GroupBox
    Friend WithEvents rbInsideL As System.Windows.Forms.RadioButton
    Friend WithEvents rbInsideM As System.Windows.Forms.RadioButton
    Friend WithEvents rbInsideH As System.Windows.Forms.RadioButton
    Friend WithEvents gbOut As System.Windows.Forms.GroupBox
    Friend WithEvents rbOutsideL As System.Windows.Forms.RadioButton
    Friend WithEvents rbOutsideM As System.Windows.Forms.RadioButton
    Friend WithEvents rbOutsideH As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rbInsideNo As System.Windows.Forms.RadioButton
    Friend WithEvents rbInsideYes As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents rbMiniNo As System.Windows.Forms.RadioButton
    Friend WithEvents rbMiniYes As System.Windows.Forms.RadioButton
    Friend WithEvents gbMiniAddress As System.Windows.Forms.GroupBox
    Friend WithEvents txtMiniAddress As System.Windows.Forms.TextBox
    Friend WithEvents gbMiniBrightness As System.Windows.Forms.GroupBox
    Friend WithEvents txtMiniBrightness As System.Windows.Forms.NumericUpDown
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents rbOutsideMH As System.Windows.Forms.RadioButton
    Friend WithEvents rbOutsideML As System.Windows.Forms.RadioButton
End Class
