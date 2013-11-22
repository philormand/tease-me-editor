<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OptionsPopUp
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
        Me.btnOk = New System.Windows.Forms.Button
        Me.cbBackup = New System.Windows.Forms.CheckBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.txtDefaltDir = New System.Windows.Forms.TextBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.tbImageFiles = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.tbAudioFiles = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbVideoFiles = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.tbLoopCheck = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.tbThumbnailSize = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.tbMaxDelay = New System.Windows.Forms.TextBox
        Me.cbNyx = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(560, 223)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 23)
        Me.btnOk.TabIndex = 0
        Me.btnOk.Text = "Ok"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'cbBackup
        '
        Me.cbBackup.AutoSize = True
        Me.cbBackup.Location = New System.Drawing.Point(19, 54)
        Me.cbBackup.Name = "cbBackup"
        Me.cbBackup.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.cbBackup.Size = New System.Drawing.Size(97, 17)
        Me.cbBackup.TabIndex = 29
        Me.cbBackup.Text = "Create Backup"
        Me.cbBackup.UseVisualStyleBackColor = True
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(16, 22)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(67, 13)
        Me.Label20.TabIndex = 28
        Me.Label20.Text = "File Location"
        '
        'txtDefaltDir
        '
        Me.txtDefaltDir.Location = New System.Drawing.Point(121, 22)
        Me.txtDefaltDir.Name = "txtDefaltDir"
        Me.txtDefaltDir.Size = New System.Drawing.Size(294, 20)
        Me.txtDefaltDir.TabIndex = 27
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(653, 223)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 30
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 145)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 32
        Me.Label1.Text = "Image Files"
        '
        'tbImageFiles
        '
        Me.tbImageFiles.Location = New System.Drawing.Point(121, 145)
        Me.tbImageFiles.Name = "tbImageFiles"
        Me.tbImageFiles.Size = New System.Drawing.Size(294, 20)
        Me.tbImageFiles.TabIndex = 31
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 188)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 13)
        Me.Label2.TabIndex = 34
        Me.Label2.Text = "Audio Files"
        '
        'tbAudioFiles
        '
        Me.tbAudioFiles.Location = New System.Drawing.Point(121, 188)
        Me.tbAudioFiles.Name = "tbAudioFiles"
        Me.tbAudioFiles.Size = New System.Drawing.Size(294, 20)
        Me.tbAudioFiles.TabIndex = 33
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 227)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(58, 13)
        Me.Label3.TabIndex = 36
        Me.Label3.Text = "Video Files"
        '
        'tbVideoFiles
        '
        Me.tbVideoFiles.Location = New System.Drawing.Point(121, 227)
        Me.tbVideoFiles.Name = "tbVideoFiles"
        Me.tbVideoFiles.Size = New System.Drawing.Size(294, 20)
        Me.tbVideoFiles.TabIndex = 35
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(15, 112)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(97, 13)
        Me.Label4.TabIndex = 38
        Me.Label4.Text = "Loop Check Depth"
        '
        'tbLoopCheck
        '
        Me.tbLoopCheck.Location = New System.Drawing.Point(121, 112)
        Me.tbLoopCheck.Name = "tbLoopCheck"
        Me.tbLoopCheck.Size = New System.Drawing.Size(67, 20)
        Me.tbLoopCheck.TabIndex = 37
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(16, 78)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(77, 13)
        Me.Label5.TabIndex = 40
        Me.Label5.Text = "Thumbnail size"
        '
        'tbThumbnailSize
        '
        Me.tbThumbnailSize.Location = New System.Drawing.Point(121, 78)
        Me.tbThumbnailSize.Name = "tbThumbnailSize"
        Me.tbThumbnailSize.Size = New System.Drawing.Size(67, 20)
        Me.tbThumbnailSize.TabIndex = 39
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(253, 112)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(57, 13)
        Me.Label6.TabIndex = 42
        Me.Label6.Text = "Max Delay"
        '
        'tbMaxDelay
        '
        Me.tbMaxDelay.Location = New System.Drawing.Point(316, 112)
        Me.tbMaxDelay.Name = "tbMaxDelay"
        Me.tbMaxDelay.Size = New System.Drawing.Size(67, 20)
        Me.tbMaxDelay.TabIndex = 41
        '
        'cbNyx
        '
        Me.cbNyx.AutoSize = True
        Me.cbNyx.Location = New System.Drawing.Point(139, 54)
        Me.cbNyx.Name = "cbNyx"
        Me.cbNyx.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.cbNyx.Size = New System.Drawing.Size(68, 17)
        Me.cbNyx.TabIndex = 43
        Me.cbNyx.Text = "Nyx Only"
        Me.cbNyx.UseVisualStyleBackColor = True
        '
        'OptionsPopUp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(752, 264)
        Me.ControlBox = False
        Me.Controls.Add(Me.cbNyx)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.tbMaxDelay)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.tbThumbnailSize)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.tbLoopCheck)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.tbVideoFiles)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbAudioFiles)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbImageFiles)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.cbBackup)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.txtDefaltDir)
        Me.Controls.Add(Me.btnOk)
        Me.Name = "OptionsPopUp"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Options"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents cbBackup As System.Windows.Forms.CheckBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents txtDefaltDir As System.Windows.Forms.TextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbImageFiles As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbAudioFiles As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbVideoFiles As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tbLoopCheck As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents tbThumbnailSize As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents tbMaxDelay As System.Windows.Forms.TextBox
    Friend WithEvents cbNyx As System.Windows.Forms.CheckBox
End Class
