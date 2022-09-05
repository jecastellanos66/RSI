<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmIntensity
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
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.gbIin = New System.Windows.Forms.GroupBox
        Me.rbIn16 = New System.Windows.Forms.RadioButton
        Me.rbIn26 = New System.Windows.Forms.RadioButton
        Me.rbIn22 = New System.Windows.Forms.RadioButton
        Me.gbOut = New System.Windows.Forms.GroupBox
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton3 = New System.Windows.Forms.RadioButton
        Me.btnClose = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        Me.gbIin.SuspendLayout()
        Me.gbOut.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackgroundImage = Global.RSI_Toteboard.My.Resources.Resources.Gradien_Light_Gray
        Me.Panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel1.Controls.Add(Me.gbIin)
        Me.Panel1.Controls.Add(Me.gbOut)
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Location = New System.Drawing.Point(4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(200, 133)
        Me.Panel1.TabIndex = 6
        '
        'gbIin
        '
        Me.gbIin.Controls.Add(Me.rbIn16)
        Me.gbIin.Controls.Add(Me.rbIn26)
        Me.gbIin.Controls.Add(Me.rbIn22)
        Me.gbIin.Location = New System.Drawing.Point(12, 56)
        Me.gbIin.Name = "gbIin"
        Me.gbIin.Size = New System.Drawing.Size(177, 39)
        Me.gbIin.TabIndex = 7
        Me.gbIin.TabStop = False
        Me.gbIin.Text = "Inside Board"
        '
        'rbIn16
        '
        Me.rbIn16.AutoSize = True
        Me.rbIn16.Location = New System.Drawing.Point(127, 19)
        Me.rbIn16.Name = "rbIn16"
        Me.rbIn16.Size = New System.Drawing.Size(45, 17)
        Me.rbIn16.TabIndex = 2
        Me.rbIn16.Tag = "3F"
        Me.rbIn16.Text = "Low"
        Me.rbIn16.UseVisualStyleBackColor = True
        '
        'rbIn26
        '
        Me.rbIn26.AutoSize = True
        Me.rbIn26.Location = New System.Drawing.Point(59, 19)
        Me.rbIn26.Name = "rbIn26"
        Me.rbIn26.Size = New System.Drawing.Size(62, 17)
        Me.rbIn26.TabIndex = 1
        Me.rbIn26.Tag = "26"
        Me.rbIn26.Text = "Medium"
        Me.rbIn26.UseVisualStyleBackColor = True
        '
        'rbIn22
        '
        Me.rbIn22.AutoSize = True
        Me.rbIn22.Location = New System.Drawing.Point(6, 19)
        Me.rbIn22.Name = "rbIn22"
        Me.rbIn22.Size = New System.Drawing.Size(47, 17)
        Me.rbIn22.TabIndex = 0
        Me.rbIn22.Tag = "22"
        Me.rbIn22.Text = "High"
        Me.rbIn22.UseVisualStyleBackColor = True
        '
        'gbOut
        '
        Me.gbOut.Controls.Add(Me.RadioButton1)
        Me.gbOut.Controls.Add(Me.RadioButton2)
        Me.gbOut.Controls.Add(Me.RadioButton3)
        Me.gbOut.Location = New System.Drawing.Point(12, 8)
        Me.gbOut.Name = "gbOut"
        Me.gbOut.Size = New System.Drawing.Size(177, 39)
        Me.gbOut.TabIndex = 10
        Me.gbOut.TabStop = False
        Me.gbOut.Text = "Outside Board"
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(127, 19)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(45, 17)
        Me.RadioButton1.TabIndex = 2
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Tag = "3F"
        Me.RadioButton1.Text = "Low"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(59, 19)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(62, 17)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Tag = "26"
        Me.RadioButton2.Text = "Medium"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(6, 19)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(47, 17)
        Me.RadioButton3.TabIndex = 0
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Tag = "22"
        Me.RadioButton3.Text = "High"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(80, 101)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(44, 23)
        Me.btnClose.TabIndex = 9
        Me.btnClose.Text = "&OK"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'frmIntensity
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.RSI_Toteboard.My.Resources.Resources.gradient
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(208, 140)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmIntensity"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Intensity"
        Me.Panel1.ResumeLayout(False)
        Me.gbIin.ResumeLayout(False)
        Me.gbIin.PerformLayout()
        Me.gbOut.ResumeLayout(False)
        Me.gbOut.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents gbIin As System.Windows.Forms.GroupBox
    Friend WithEvents rbIn22 As System.Windows.Forms.RadioButton
    Friend WithEvents rbIn16 As System.Windows.Forms.RadioButton
    Friend WithEvents rbIn26 As System.Windows.Forms.RadioButton
    Friend WithEvents gbOut As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
End Class
