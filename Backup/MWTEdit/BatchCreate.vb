Public Class BatchCreate

    Private Sub btnBCOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBCOK.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub BatchCreate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.AcceptButton = btnBCOK
        Me.CancelButton = btnBCCanc
    End Sub

    Private Sub btnBCCanc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBCCanc.Click
        Me.Close()
    End Sub
End Class