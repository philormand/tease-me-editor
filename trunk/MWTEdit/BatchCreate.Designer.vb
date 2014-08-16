<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BatchCreate
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
        Me.btnBCOK = New System.Windows.Forms.Button
        Me.btnBCCanc = New System.Windows.Forms.Button
        Me.txtPrefix = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.cbxContinue = New System.Windows.Forms.CheckBox
        Me.cbxDelay = New System.Windows.Forms.CheckBox
        Me.txtFrom = New System.Windows.Forms.MaskedTextBox
        Me.txtTo = New System.Windows.Forms.MaskedTextBox
        Me.txtDelay = New System.Windows.Forms.MaskedTextBox
        Me.SuspendLayout()
        '
        'btnBCOK
        '
        Me.btnBCOK.Location = New System.Drawing.Point(257, 227)
        Me.btnBCOK.Name = "btnBCOK"
        Me.btnBCOK.Size = New System.Drawing.Size(75, 23)
        Me.btnBCOK.TabIndex = 0
        Me.btnBCOK.Text = "Ok"
        Me.btnBCOK.UseVisualStyleBackColor = True
        '
        'btnBCCanc
        '
        Me.btnBCCanc.Location = New System.Drawing.Point(348, 227)
        Me.btnBCCanc.Name = "btnBCCanc"
        Me.btnBCCanc.Size = New System.Drawing.Size(75, 23)
        Me.btnBCCanc.TabIndex = 1
        Me.btnBCCanc.Text = "Cancel"
        Me.btnBCCanc.UseVisualStyleBackColor = True
        '
        'txtPrefix
        '
        Me.txtPrefix.Location = New System.Drawing.Point(96, 12)
        Me.txtPrefix.Name = "txtPrefix"
        Me.txtPrefix.Size = New System.Drawing.Size(100, 20)
        Me.txtPrefix.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Page Prefix"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "From Number"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(177, 52)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "To Number"
        '
        'cbxContinue
        '
        Me.cbxContinue.AutoSize = True
        Me.cbxContinue.Location = New System.Drawing.Point(19, 93)
        Me.cbxContinue.Name = "cbxContinue"
        Me.cbxContinue.Size = New System.Drawing.Size(102, 17)
        Me.cbxContinue.TabIndex = 6
        Me.cbxContinue.Text = "Continue Button"
        Me.cbxContinue.UseVisualStyleBackColor = True
        '
        'cbxDelay
        '
        Me.cbxDelay.AutoSize = True
        Me.cbxDelay.Location = New System.Drawing.Point(19, 127)
        Me.cbxDelay.Name = "cbxDelay"
        Me.cbxDelay.Size = New System.Drawing.Size(53, 17)
        Me.cbxDelay.TabIndex = 7
        Me.cbxDelay.Text = "Delay"
        Me.cbxDelay.UseVisualStyleBackColor = True
        '
        'txtFrom
        '
        Me.txtFrom.Location = New System.Drawing.Point(96, 44)
        Me.txtFrom.Mask = "00000"
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.Size = New System.Drawing.Size(54, 20)
        Me.txtFrom.TabIndex = 8
        Me.txtFrom.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtFrom.ValidatingType = GetType(Integer)
        '
        'txtTo
        '
        Me.txtTo.Location = New System.Drawing.Point(243, 49)
        Me.txtTo.Mask = "00000"
        Me.txtTo.Name = "txtTo"
        Me.txtTo.Size = New System.Drawing.Size(54, 20)
        Me.txtTo.TabIndex = 9
        Me.txtTo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtTo.ValidatingType = GetType(Integer)
        '
        'txtDelay
        '
        Me.txtDelay.Location = New System.Drawing.Point(96, 124)
        Me.txtDelay.Mask = "00000"
        Me.txtDelay.Name = "txtDelay"
        Me.txtDelay.Size = New System.Drawing.Size(54, 20)
        Me.txtDelay.TabIndex = 10
        Me.txtDelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtDelay.ValidatingType = GetType(Integer)
        '
        'BatchCreate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(442, 262)
        Me.ControlBox = False
        Me.Controls.Add(Me.txtDelay)
        Me.Controls.Add(Me.txtTo)
        Me.Controls.Add(Me.txtFrom)
        Me.Controls.Add(Me.cbxDelay)
        Me.Controls.Add(Me.cbxContinue)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtPrefix)
        Me.Controls.Add(Me.btnBCCanc)
        Me.Controls.Add(Me.btnBCOK)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "BatchCreate"
        Me.Text = "BatchCreate"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnBCOK As System.Windows.Forms.Button
    Friend WithEvents btnBCCanc As System.Windows.Forms.Button
    Friend WithEvents txtPrefix As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbxContinue As System.Windows.Forms.CheckBox
    Friend WithEvents cbxDelay As System.Windows.Forms.CheckBox
    Friend WithEvents txtFrom As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtTo As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtDelay As System.Windows.Forms.MaskedTextBox
End Class
