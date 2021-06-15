Public Class Form1
    Dim Mk1, Mk2, Mk3, Mk4, Mk5, Mk6 As Single 'Mk = Mata Kuliah rentang nilai 1-100
    Dim n1, n2, n3, n4, n5, n6 As Single 'n = Nilai konversi dari 1-100 ke 1-4
    Dim s1, s2, s3, s4, s5, s6 As Single 's = Bobot SKS
    Dim ip As Single 'ip = IPK
    Dim sliding As String = "close"
    Dim H1E020056, H1E020002, H1E020012, H1E020042 As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Mk1 = TextBox1.Text
        Select Case Val(Mk1)
            Case 80 To 100
                n1 = 4
            Case 70 To 79
                n1 = 3
            Case 60 To 69
                n1 = 2
            Case 50 To 59
                n1 = 1
            Case <= 49
                n1 = 0
        End Select
        Mk2 = TextBox2.Text
        Select Case Val(Mk2)
            Case 80 To 100
                n2 = 4
            Case 70 To 79
                n2 = 3
            Case 60 To 69
                n2 = 2
            Case 50 To 59
                n2 = 1
            Case <= 49
                n2 = 0
        End Select
        Mk3 = TextBox3.Text
        Select Case Val(Mk3)
            Case 80 To 100
                n3 = 4
            Case 70 To 79
                n3 = 3
            Case 60 To 69
                n3 = 2
            Case 50 To 59
                n3 = 1
            Case <= 49
                n3 = 0
        End Select
        Mk4 = TextBox4.Text
        Select Case Val(Mk4)
            Case 80 To 100
                n4 = 4
            Case 70 To 79
                n4 = 3
            Case 60 To 69
                n4 = 2
            Case 50 To 59
                n4 = 1
            Case <= 49
                n4 = 0
        End Select
        Mk5 = TextBox5.Text
        Select Case Val(Mk5)
            Case 80 To 100
                n5 = 4
            Case 70 To 79
                n5 = 3
            Case 60 To 69
                n5 = 2
            Case 50 To 59
                n5 = 1
            Case <= 49
                n5 = 0
        End Select
        Mk6 = TextBox6.Text
        Select Case Val(Mk6)
            Case 80 To 100
                n6 = 4
            Case 70 To 79
                n6 = 3
            Case 60 To 69
                n6 = 2
            Case 50 To 59
                n6 = 1
            Case <= 49
                n6 = 0
        End Select
        s1 = TextBox7.Text
        s2 = TextBox8.Text
        s3 = TextBox9.Text
        s4 = TextBox10.Text
        s5 = TextBox11.Text
        s6 = TextBox12.Text
        ip = htg.cnt.count(n1, n2, n3, n4, n5, n6, s1, s2, s3, s4, s5, s6)
        TextBox13.Text = ip
        TextBox13.Text = FormatNumber(TextBox13.Text, 2)
        Select Case Val(ip)
            Case 4
                ket.Text = "Cieee Cumlaude"
            Case 3.5 To 3.9
                ket.Text = "Idaman Mertua"
            Case 3 To 3.4
                ket.Text = "Fix Kamu Keren!"
            Case 2.5 To 2.9
                ket.Text = "Tingkatkan Yaa!"
            Case 2 To 2.5
                ket.Text = "Ngapain Aja Si!"
            Case 1 To 1.9
                ket.Text = " Serius Dong!!!"
            Case < 0.9
                ket.Text = "Lah di Drop Out"
        End Select
    End Sub

    Private Sub btnsld_Click(sender As Object, e As EventArgs) Handles btnsld.Click
        Timer1.Start()
        Timer2.Start()

    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If sliding = "open" Then
            sld.Width += 5
            If sld.Width >= 145 Then
                Timer1.Stop()
                nama.Visible = True
                nim.Visible = True
                prodi.Visible = True
                sliding = "close"

            End If
        Else
            sld.Width -= 5
            If sld.Width <= 40 Then
                Timer1.Stop()
                nama.Visible = False
                nim.Visible = False
                prodi.Visible = False
                sliding = "open"
            End If
        End If

    End Sub
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If pnl.Location = New Point(160, 50) Then
            pnl.Location = New Point(95, 50)
            Timer2.Stop()
        ElseIf pnl.Location = New Point(95, 50) Then
            pnl.Location = New Point(160, 50)
            Timer2.Stop()
        End If

    End Sub

    '===================================================================
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Close()
    End Sub

    Private Sub lgn_Click(sender As Object, e As EventArgs) Handles lgn.Click
        nama.Text = lgnnama.Text
        nim.Text = lgnnim.Text
        prodi.Text = lgnprodi.Text
        login.Visible = False
        If lgnnim.Text = H1E020002 Then
            PictureBox2.Visible = True
            PictureBox3.Visible = False
            PictureBox4.Visible = False
            PictureBox5.Visible = False
        ElseIf lgnnim.Text = H1E020012 Then
            PictureBox3.Visible = True
            PictureBox2.Visible = False
            PictureBox4.Visible = False
            PictureBox5.Visible = False
        ElseIf lgnnim.Text = H1E020042 Then
            PictureBox5.Visible = True
            PictureBox2.Visible = False
            PictureBox3.Visible = False
            PictureBox4.Visible = False
        ElseIf lgnnim.Text = H1E020056 Then
            PictureBox4.Visible = True
            PictureBox2.Visible = False
            PictureBox3.Visible = False
            PictureBox5.Visible = False
        End If

    End Sub
    Private Sub button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Timer3.Start()

    End Sub
    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        If sliding = "open" Then
            pnl.Height += 5
            If pnl.Height >= 265 Then
                Timer3.Stop()
                sliding = "close"

            End If
        Else
            pnl.Height -= 5
            If pnl.Height <= 40 Then
                Timer3.Stop()
                sliding = "open"
            End If
        End If
    End Sub
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        login.Visible = True
        lgnnama.Text = ""
        lgnnim.Text = ""
        lgnprodi.Text = ""
    End Sub
    Private Sub button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
    End Sub
End Class
