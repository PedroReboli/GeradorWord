Module avarias
    Public Function avaria()
        acesso.modi("ssresponsavel", Form1.TextBox24.Text, True)
        acesso.modi("ssSEGURADO", Form1.TextBox2.Text, True)
        acesso.modi("sstaq", Form1.TextBox1.Text, True)
        acesso.modi("ssteq", Form1.TextBox17.Text, True)
        acesso.modi("sscorretora", Form1.TextBox3.Text, True)
        acesso.modi("ssemail", Form1.TextBox25.Text, True)
        acesso.modi("sstelefone", Form1.TextBox26.Text, True)
        acesso.modi("sscnpj", Form1.TextBox28.Text, True)
        acesso.modi("ssendereco", Form1.TextBox27.Text, True)
        acesso.modi("ssmercadorias", Form1.TextBox4.Text, True)
        acesso.modi("sslmg", Form1.TextBox5.Text + ",00", True)
        acesso.modi("ssextlmg", Form1.TextBox6.Text, True)
        acesso.modi("ssavaria", Form1.TextBox12.Text + ",00", True)
        acesso.modi("ssextavaria", Form1.TextBox13.Text, True)
        acesso.modi("ssdesc", Form1.TextBox7.Text, True)
        acesso.modi("sstaxa", Form1.TextBox16.Text, True)
        'acesso.modi("ssfranq", Form1.TextBox14.Text + ",00", True)
        acesso.modi("ssfranq", Form1.TextBox14.Text, True)
        Dim temp As String
        temp = Form1.TextBox14.Text.ToString.Replace(".", "")
        If temp = "" Then
            temp = "0"
        End If
        temp = extensu.NumeroToExtenso(temp)
        temp = temp.Replace("Reais", "")
        acesso.modi("ssextfranq", temp, True) 'adicionar extenso
        acesso.modi("sslmo", Form1.TextBox15.Text + ",00", True)
        acesso.modi("ssextlmo", extensu.NumeroToExtenso(Form1.TextBox15.Text.Replace(".", "")), True)
        acesso.modi("sspremio", Form1.TextBox8.Text + ",00", True)
        acesso.modi("ssextpremio", Form1.TextBox29.Text, True)
        acesso.modi("ssdate1", Form1.TextBox9.Text, True)
        acesso.modi("ssdata2", Form1.TextBox10.Text, True)
        acesso.modi("ssdata3", Form1.TextBox11.Text, True)
    End Function
End Module
