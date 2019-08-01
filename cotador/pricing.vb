Imports Microsoft.Office.Interop
Imports System.Globalization
Imports System.IO
Module pricing
    Public Function calcular(ByVal path2)
        Dim appXL As Excel.Application
        Dim wbXL As Excel.Workbook
        Dim wbsXL As Excel.Workbooks
        Dim shXL As Excel.Worksheet
        Dim Checker As Integer

        'appXL = CreateObject("excel.application")
        'appXL.Visible = False
        'AAAAAABABBA;
        'wbsXL = appXL.Workbooks
        'wbXL = wbsXL.Open(path2)
        'shXL = appXL.Worksheets(2)
        If System.IO.File.ReadAllLines(Path.GetTempPath + "\saida.txt")(0) = "Sim" Then

            Form1.CheckBox4.Checked = True
        Else
            Form1.CheckBox4.Checked = False
        End If 'roubo
        Dim PM As Double
        PM = System.IO.File.ReadAllLines(Path.GetTempPath + "\saida.txt")(1) 'premio minimo
        Form1.TextBox8.Text = PM.ToString("C", CultureInfo.CreateSpecificCulture("en-US")).Replace("$", "")
        If System.IO.File.ReadAllLines(Path.GetTempPath + "\saida.txt")(2) > 0 Then 'avaria
            Form1.CheckBox3.Checked = True
        Else
            Form1.CheckBox3.Checked = False
        End If 'avaria
        Dim LMG As Double
        Dim desconto
        LMG = System.IO.File.ReadAllLines(Path.GetTempPath + "\saida.txt")(3) 'LMG
        Form1.TextBox5.Text = LMG.ToString("C", CultureInfo.CreateSpecificCulture("en-US")).Replace("$", "")
        Form1.TextBox7.Text = System.IO.File.ReadAllLines(Path.GetTempPath + "\saida.txt")(4) * 100 'desconto
        Form1.TextBox20.Text = System.IO.File.ReadAllLines(Path.GetTempPath + "\saida.txt")(5) ' taxa roubo
        Form1.TextBox17.Text = System.IO.File.ReadAllLines(Path.GetTempPath + "\saida.txt")(6) 'TAQ RCF-DC
        Form1.TextBox1.Text = System.IO.File.ReadAllLines(Path.GetTempPath + "\saida.txt")(7) 'TAQ RCTR-C
        '\\\\\\\\\\\\\\\\\\\\ parte 2
        'shXL = appXL.Worksheets(1)
        Form1.TextBox27.Text = System.IO.File.ReadAllLines(Path.GetTempPath + "\saida.txt")(8) 'endereco
        Form1.TextBox3.Text = System.IO.File.ReadAllLines(Path.GetTempPath + "\saida.txt")(9) 'corretora
        Form1.TextBox2.Text = System.IO.File.ReadAllLines(Path.GetTempPath + "\saida.txt")(10) 'proponente
        Form1.TextBox28.Text = System.IO.File.ReadAllLines(Path.GetTempPath + "\saida.txt")(11) 'CNPJ
        'System.Threading.Thread.Sleep(5000)
        'appXL.ActiveWorkbook.Close(False)
        'wbsXL.Close()
        'appXL.Quit()
    End Function
End Module
