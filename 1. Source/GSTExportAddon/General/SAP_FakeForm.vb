Public Class SAP_FakeForm

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click

        If Me.WindowState = FormWindowState.Maximized Then
            Me.WindowState = FormWindowState.Normal
        Else
            Me.WindowState = FormWindowState.Maximized
        End If
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        Me.Close()
    End Sub

    Private Sub Panel1_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel1.DoubleClick
        
    End Sub

    Private Sub Panel1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseDown
        If e.Clicks = 2 Then
            If Me.WindowState = FormWindowState.Maximized Then
                Me.WindowState = FormWindowState.Normal
            ElseIf Me.WindowState = FormWindowState.Minimized Then
                Me.WindowState = FormWindowState.Normal
            Else
                Me.WindowState = FormWindowState.Maximized
            End If
        End If
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Panel1.Capture = False

            ' Create and send a WM_NCLBUTTONDOWN message.
            Const WM_NCLBUTTONDOWN As Integer = &HA1S
            Const HTCAPTION As Integer = 2
            Dim msg As Message = _
                Message.Create(Me.Handle, WM_NCLBUTTONDOWN, _
                    New IntPtr(HTCAPTION), IntPtr.Zero)
            Me.DefWndProc(msg)
        End If

    End Sub

    Private Sub SAP_FakeForm_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Application.DoEvents()
        Me.BringToFront()
        Me.TopMost = True
    End Sub

  
    Private Sub SAP_FakeForm_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        ' WindowState = FormWindowState.Maximized
    End Sub
End Class