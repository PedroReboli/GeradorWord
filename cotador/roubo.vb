Module roubo
    Public Function roubo()
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
        acesso.modi("sstaxa", Form1.TextBox20.Text, True)
        acesso.modi("sspos1", Form1.TextBox21.Text, True)
        acesso.modi("sspos2", Form1.TextBox22.Text, True)
        acesso.modi("sspos3", Form1.TextBox23.Text, True)
        acesso.modi("ssextpos1", extensu.NumeroToExtenso(Form1.TextBox21.Text).Replace(" Reais", ""), True)
        acesso.modi("ssextpos2", extensu.NumeroToExtenso(Form1.TextBox22.Text).Replace(" Reais", ""), True)
        acesso.modi("ssextpos3", extensu.NumeroToExtenso(Form1.TextBox23.Text).Replace(" Reais", ""), True)
        'adicionar extenso para POS
        '
        '
        acesso.modi("sspremio", Form1.TextBox8.Text + ",00", True)
        acesso.modi("ssextpremio", Form1.TextBox29.Text, True)

        acesso.modi("ssdate1", Form1.TextBox9.Text, True)
        acesso.modi("ssdata2", Form1.TextBox10.Text, True)
        acesso.modi("ssdata3", Form1.TextBox11.Text, True)



    End Function
End Module
