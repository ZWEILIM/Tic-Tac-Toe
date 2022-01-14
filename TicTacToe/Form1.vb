Public Class Form1


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For Each c As Control In Panel2.Controls

            If c.GetType = GetType(Button) Then
                AddHandler c.Click, AddressOf btn_Click

            End If
        Next
    End Sub

    Dim controlXO As Integer = 0
    Private Sub btn_Click(sender As Object, e As EventArgs)
        Dim btn As Button = sender

        If btn.Text.Equals("") Then

            If controlXO Mod 2 = 0 Then
                btn.Text = "X"
                btn.ForeColor = Color.Red
                Label1.Text = "O turn"
                getWinner()
            Else
                btn.Text = "O"
                btn.ForeColor = Color.Blue
                Label1.Text = "X turn"
                getWinner()
            End If
            controlXO += 1

        End If

    End Sub

    Dim win As Boolean = False

    Private Sub getWinner()
        If Not Button1.Text.Equals("") AndAlso Button1.Text.Equals(Button2.Text) AndAlso Button1.Text.Equals(Button3.Text) Then
            win = True
            wineffect(Button1, Button2, Button3)
            Form2.Show()
        End If

        If Not Button1.Text.Equals("") AndAlso Button1.Text.Equals(Button4.Text) AndAlso Button1.Text.Equals(Button7.Text) Then
            win = True
            wineffect(Button1, Button4, Button7)
            Form2.Show()
        End If
        If Not Button1.Text.Equals("") AndAlso Button1.Text.Equals(Button5.Text) AndAlso Button1.Text.Equals(Button9.Text) Then
            win = True
            wineffect(Button1, Button5, Button9)
            Form2.Show()
        End If
        If Not Button2.Text.Equals("") AndAlso Button2.Text.Equals(Button5.Text) AndAlso Button2.Text.Equals(Button8.Text) Then
            win = True
            wineffect(Button2, Button5, Button8)
            Form2.Show()
        End If
        If Not Button3.Text.Equals("") AndAlso Button3.Text.Equals(Button6.Text) AndAlso Button3.Text.Equals(Button9.Text) Then
            win = True
            wineffect(Button3, Button6, Button9)
            Form2.Show()
        End If
        If Not Button4.Text.Equals("") AndAlso Button4.Text.Equals(Button5.Text) AndAlso Button4.Text.Equals(Button6.Text) Then
            win = True
            wineffect(Button4, Button5, Button6)
            Form2.Show()
        End If
        If Not Button7.Text.Equals("") AndAlso Button7.Text.Equals(Button8.Text) AndAlso Button7.Text.Equals(Button9.Text) Then
            win = True
            wineffect(Button7, Button8, Button9)
            Form2.Show()
        End If
        If Not Button7.Text.Equals("") AndAlso Button7.Text.Equals(Button5.Text) AndAlso Button7.Text.Equals(Button3.Text) Then
            win = True
            wineffect(Button7, Button5, Button3)
            Form2.Show()
        End If

        If btntextslength() = 9 AndAlso win = False Then
            Label1.Text = "No Winner"
        End If
    End Sub

    Function btntextslength() As Integer

        Dim btnl As Integer = 0
        For Each c As Control In Panel2.Controls

            If c.GetType = GetType(Button) Then
                btnl += c.Text.Length

            End If
        Next
        Return btnl
    End Function

    Private Sub wineffect(ByVal b1 As Button, ByVal b2 As Button, ByVal b3 As Button)

        b1.BackColor = Color.Orange
        b2.BackColor = Color.Orange
        b3.BackColor = Color.Orange

        b1.ForeColor = Color.White
        b2.ForeColor = Color.White
        b3.ForeColor = Color.White

        Label1.Text = b1.Text + " Win"

    End Sub

    Private Sub ButtonNewGame_Click(sender As Object, e As EventArgs) Handles ButtonNewGame.Click

        controlXO = 0
        win = False
        Label1.Text = "Play"
        For Each c As Control In Panel2.Controls

            If c.GetType = GetType(Button) Then
                c.BackColor = Color.White
                c.Text = ""
            End If
        Next
    End Sub
End Class
