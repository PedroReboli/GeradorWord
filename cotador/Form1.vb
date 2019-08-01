Imports Microsoft.Office.Interop

Public Class Form1
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            Label11.Visible = True
            Label12.Visible = True
            Label13.Visible = True
            Label14.Visible = True
            Label15.Visible = True
            Label17.Visible = True
            TextBox12.Visible = True
            TextBox13.Visible = True
            TextBox14.Visible = True
            TextBox15.Visible = True
            TextBox16.Visible = True
            TextBox18.Visible = True

        End If
        If CheckBox3.Checked = False Then
            Label11.Visible = False
            Label12.Visible = False
            Label13.Visible = False
            Label14.Visible = False
            Label15.Visible = False
            Label17.Visible = False
            TextBox12.Visible = False
            TextBox13.Visible = False
            TextBox14.Visible = False
            TextBox15.Visible = False
            TextBox16.Visible = False
            TextBox18.Visible = False
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            Label16.Visible = True
            Label19.Visible = True
            Label20.Visible = True
            TextBox17.Visible = True
            TextBox20.Visible = True
            TextBox21.Visible = True
            TextBox22.Visible = True
            TextBox23.Visible = True
        End If
        If CheckBox4.Checked = False Then
            Label16.Visible = False
            Label19.Visible = False
            Label20.Visible = False
            TextBox17.Visible = False
            TextBox20.Visible = False
            TextBox21.Visible = False
            TextBox22.Visible = False
            TextBox23.Visible = False

        End If
    End Sub
    Public path As String
    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Try
            TextBox6.Text = extensu.NumeroToExtenso(TextBox5.Text.Replace(",00", "").Replace(".", ""))
        Catch ex As Exception

        End Try


    End Sub

    Public oWord As Word.Application
    Public oDoc As Word.Document
    Public objDoc As Word.Document
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        If CheckBox3.Checked = True Then
            oDoc = oWord.Documents.Open(path + "acidente.docx")
            objDoc = oDoc
            avarias.avaria()
            If CheckBox4.Checked = True Then
                oDoc = oWord.Documents.Open(path + "roubo.docx")
                objDoc = oDoc
                roubo.roubo()
            End If
        ElseIf CheckBox4.Checked = True Then
            oDoc = oWord.Documents.Open(path + "roubo.docx")
            objDoc = oDoc
            roubo.roubo()
            oDoc = oWord.Documents.Open(path + "avaria.docx")
            objDoc = oDoc
            padrao.padrao()
        Else
            oDoc = oWord.Documents.Open(path + "avaria.docx")
            objDoc = oDoc
            padrao.padrao()
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim dialog As New OpenFileDialog
        dialog.Filter = "|check.txt"
        If dialog.ShowDialog() = DialogResult.OK Then
            Button2.BackColor = Color.Green
            path = dialog.FileName.Replace("check.txt", "")
            Button1.Enabled = True
            My.Settings.path = path
            My.Settings.Save()
        End If
    End Sub
    Private Sub Form1_DragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        Dim path3 As String
        For Each path3 In files
            pricing.calcular(path3)
        Next
    End Sub
    Private Sub Form1_DragEnter(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.AllowDrop = True
        If Not My.Settings.path = "" And System.IO.File.Exists(My.Settings.path + "check.txt") Then
            Button2.BackColor = Color.Green
            Button1.Enabled = True
            path = My.Settings.path
        End If
        Dim today As System.DateTime
        Dim answer As System.DateTime
        today = System.DateTime.Now
        answer = today.AddDays(30)
        TextBox9.Text = answer.Day
        TextBox10.Text = answer.Month
        TextBox11.Text = answer.Year

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        Try
            TextBox29.Text = extensu.NumeroToExtenso(TextBox8.Text.Replace(",00", "").Replace(".", ""))
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        Try
            TextBox13.Text = extensu.NumeroToExtenso(TextBox12.Text.Replace(",00", "").Replace(".", ""))
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        Try
            TextBox18.Text = extensu.NumeroToExtenso(TextBox15.Text.Replace(",00", "").Replace(".", ""))
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Controls.Clear() 'removes all the controls on the form
        InitializeComponent() 'load all the controls again
        Form1_Load(e, e)
    End Sub
End Class
