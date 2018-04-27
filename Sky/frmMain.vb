
Public Class frmMain
    Dim threadNYSE As System.Threading.Thread
    Dim threadNASDAQ As System.Threading.Thread
    Dim threadNYSEAMEX As System.Threading.Thread
    Dim threadPINK As System.Threading.Thread
    Dim threadTSE As System.Threading.Thread
    Dim threadCycleRun As System.Threading.Thread
    Dim threadTotalMin As System.Threading.Thread
    Dim threadCicle60Min As System.Threading.Thread


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
        DataCounter = 0


    End Sub

    Private Sub mainCodeTotalOk()
        Dim Run As Integer = 0
        Do Until Run = 1
            txtBoxTotalOK.Text = TotalOk
        Loop
    End Sub
    Private Sub mainCodeTotalNOk()
        Dim Run As Integer = 0
        Do Until Run = 1
            txtBoxTotalOK.Text = TotalOk
        Loop
    End Sub
    Private Sub mainCodeTotalOkNOK()
        Dim Run As Integer = 0
        Do Until Run = 1
            txtBoxTotalOK.Text = TotalOk
        Loop
    End Sub

    ' Private Sub ProgressBar()

    ' Dim CicleRun As Integer = 0
    '  Do Until CicleRun = 1
    '   If ProgressBarValue >= 20 Then
    '      ProgressBarValue = 0
    ' End If
    'ProgressBar1.Value = ProgressBarValue
    ' Loop

    'End Sub

    Private Sub inThreadMainCodeExtractionNYSE()
        Try
            Dim Alfabeto() As String = {"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
            Dim Letra01, Letra02, Letra03, Letra04, Letra05 As String
            Dim Total, nok As Integer
            Dim NYSERunOk As Integer = 0
            nok = 1
            Total = 1
            Dim SourceCode As String = ""
            txtMarket01.Text = ("NYSE")
            For Each Letra05 In Alfabeto
                For Each Letra04 In Alfabeto
                    For Each Letra03 In Alfabeto
                        For Each Letra02 In Alfabeto
                            For Each Letra01 In Alfabeto

                                '____________________________________
                                If SkyRun = 1 Then
                                    '____________________________________

                                    txtLetter05_01.Text = Letra05
                                    txtLetter04_01.Text = Letra04
                                    txtLetter03_01.Text = Letra03
                                    txtLetter02_01.Text = Letra02
                                    txtLetter01_01.Text = Letra01
                                    Dim symbol As String = (txtLetter05_01.Text) + (txtLetter04_01.Text) + (txtLetter03_01.Text) + (txtLetter02_01.Text) + (txtLetter01_01.Text)
                                    'Navegar até...
                                    'WebBrowser1.Navigate(txtURL.Text)
                                    'Sacar SourceCode
                                    'Implementar aqui verificação de erro para ligação á internet
                                    Try
                                        Dim Request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create("http://finance.google.com/finance?q=" + txtMarket01.Text + "%3A" + symbol + "&fstype=ii&hl=en")
                                        Dim Response As System.Net.HttpWebResponse = Request.GetResponse()
                                        Dim sr As System.IO.StreamReader = New System.IO.StreamReader(Response.GetResponseStream())
                                        SourceCode = sr.ReadToEnd()
                                    Catch ex As Exception
                                        NYSERunOk = 0
                                        Do Until NYSERunOk = 1
                                            Try
                                                txtCompany01.Text = ("Connecting...")
                                                Dim Request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create("http://finance.google.com/finance?q=" + txtMarket01.Text + "%3A" + symbol + "&fstype=ii&hl=en")
                                                Dim Response As System.Net.HttpWebResponse = Request.GetResponse()
                                                Dim sr As System.IO.StreamReader = New System.IO.StreamReader(Response.GetResponseStream())
                                                SourceCode = sr.ReadToEnd()
                                                NYSERunOk = 1
                                            Catch ez As Exception
                                                NYSERunOk = 0
                                                txtCompany01.Text = ("Connecting...")
                                                txtFail01.Text += 1
                                                System.Threading.Thread.Sleep(60000)
                                            End Try
                                        Loop
                                    End Try

                                    ProgressBarValue = ProgressBarValue + 1

                                    SourceCode = CharRemover(SourceCode)
                                    'GRAVAR ficheiro com fonte
                                    ' Dim FILE_NAME As String = "C:\Users\Rui\Desktop\temp.txt"
                                    ' If System.IO.File.Exists(FILE_NAME) = True Then
                                    'Dim objWriter As New System.IO.StreamWriter(FILE_NAME)
                                    ' objWriter.Write(SourceCode)
                                    'objWriter.Close()
                                    'Else
                                    ' MsgBox("File does not exist")
                                    ' End If
                                    'VERIFICAR se ficheiro ok para extrair dados
                                    Dim test01, test02 As Integer
                                    test01 = 0
                                    test02 = 0
                                    txtCompany01.Text = ("")
                                    test01 = InStr(SourceCode, "<div class=g-section sfe-break-bottom-16 overflow-floatfix>")
                                    test02 = InStr(SourceCode, "<td class=lft lm>Revenue")
                                    If test01 <> 0 And test02 <> 0 Then
                                        Dim strTFind, strTInicial, strTFinal, resT As String
                                        Dim tamTFind, tamTInicial, tamTFinal, tamTDesc As Integer
                                        Dim posTFind, posTInicial, posTFinal As Integer
                                        resT = ""
                                        strTFind = ("<div class=g-section sfe-break-bottom-16 overflow-floatfix>")
                                        strTInicial = ("<h3>")
                                        strTFinal = ("</h3>")
                                        tamTFind = Len(strTFind)
                                        tamTInicial = Len(strTInicial)
                                        tamTFinal = Len(strTFinal)
                                        posTFind = InStr(SourceCode, strTFind)
                                        posTFind = posTFind + tamTFind
                                        posTInicial = InStr(posTFind, SourceCode, strTInicial)
                                        posTInicial = posTInicial + tamTInicial
                                        posTFinal = InStr(posTInicial, SourceCode, strTFinal)
                                        tamTDesc = posTFinal - posTInicial
                                        resT = Mid(SourceCode, posTInicial, tamTDesc)
                                        txtCompany01.Text = resT

                                        lstBox.Items.Insert(0, txtMarket01.Text + " - " + symbol + " ---- " + txtCompany01.Text)
                                        lstBoxDATA.Items.Add(txtMarket01.Text + " - " + symbol + " ---- " + txtCompany01.Text)

                                        txtBoxData.Text += (txtMarket01.Text + " - " + symbol + " ---- " + txtCompany01.Text)









                                        IncomeStatment(SourceCode)
                                        BalanceSheet(SourceCode)
                                        CashFlow(SourceCode)

                                        'Insert in DATABASE here

                                        txtOK01.Text = Total
                                        Total = Total + 1
                                        TotalOk = TotalOk + 1
                                        txtBoxTotalOK.Text = TotalOk
                                    Else
                                        'MsgBox("Ficheiro Não Válido")
                                        txtNOK01.Text = nok
                                        nok = nok + 1
                                        TotalNOk = TotalNOk + 1
                                        txtBoxTotalNOK.Text = TotalNOk
                                    End If

                                    '____________________________________
                                Else
                                    Do Until SkyRun = 1
                                        txtStatus01.Text = "Paused"
                                        System.Threading.Thread.Sleep(100)
                                    Loop
                                End If
                                txtStatus01.Text = "Running"
                                '____________________________________

                            Next
                        Next
                    Next
                Next
            Next
        Catch ex As Exception
            txtStatus01.Text = ("ERROR")
            btnRefresh01.Visible = True
        End Try
    End Sub
    Private Sub inThreadMainCodeExtractionNASDAQ()
        Try
            Dim Alfabeto() As String = {"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
            Dim Letra01, Letra02, Letra03, Letra04, Letra05 As String
            Dim Total, nok As Integer
            Dim NASDAQRunOk As Integer = 0
            nok = 1
            Total = 1
            Dim SourceCode As String = ""
            txtMarket02.Text = ("NASDAQ")
            For Each Letra05 In Alfabeto
                For Each Letra04 In Alfabeto
                    For Each Letra03 In Alfabeto
                        For Each Letra02 In Alfabeto
                            For Each Letra01 In Alfabeto

                                '____________________________________
                                If SkyRun = 1 Then
                                    '____________________________________

                                    txtLetter05_02.Text = Letra05
                                    txtLetter04_02.Text = Letra04
                                    txtLetter03_02.Text = Letra03
                                    txtLetter02_02.Text = Letra02
                                    txtLetter01_02.Text = Letra01
                                    Dim symbol As String = (txtLetter05_02.Text) + (txtLetter04_02.Text) + (txtLetter03_02.Text) + (txtLetter02_02.Text) + (txtLetter01_02.Text)

                                    'Navegar até...
                                    'WebBrowser1.Navigate(txtURL.Text)
                                    'Sacar SourceCode
                                    'Implementar aqui verificação de erro para ligação á internet
                                    Try
                                        Dim Request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create("http://finance.google.com/finance?q=" + txtMarket02.Text + "%3A" + symbol + "&fstype=ii&hl=en")
                                        Dim Response As System.Net.HttpWebResponse = Request.GetResponse()
                                        Dim sr As System.IO.StreamReader = New System.IO.StreamReader(Response.GetResponseStream())
                                        SourceCode = sr.ReadToEnd()
                                    Catch ex As Exception
                                        NASDAQRunOk = 0
                                        Do Until NASDAQRunOk = 1
                                            Try
                                                txtCompany02.Text = ("Connecting...")
                                                Dim Request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create("http://finance.google.com/finance?q=" + txtMarket02.Text + "%3A" + symbol + "&fstype=ii&hl=en")
                                                Dim Response As System.Net.HttpWebResponse = Request.GetResponse()
                                                Dim sr As System.IO.StreamReader = New System.IO.StreamReader(Response.GetResponseStream())
                                                SourceCode = sr.ReadToEnd()
                                                NASDAQRunOk = 1
                                            Catch ez As Exception
                                                NASDAQRunOk = 0
                                                txtCompany02.Text = ("Connecting...")
                                                txtFail02.Text += 1
                                                System.Threading.Thread.Sleep(60000)
                                            End Try
                                        Loop
                                    End Try

                                    ProgressBarValue = ProgressBarValue + 1

                                    SourceCode = CharRemover(SourceCode)
                                    'GRAVAR ficheiro com fonte
                                    ' Dim FILE_NAME As String = "C:\Users\Rui\Desktop\temp.txt"
                                    ' If System.IO.File.Exists(FILE_NAME) = True Then
                                    'Dim objWriter As New System.IO.StreamWriter(FILE_NAME)
                                    ' objWriter.Write(SourceCode)
                                    'objWriter.Close()
                                    'Else
                                    ' MsgBox("File does not exist")
                                    ' End If
                                    'VERIFICAR se ficheiro ok para extrair dados
                                    Dim test01, test02 As Integer
                                    test01 = 0
                                    test02 = 0
                                    txtCompany02.Text = ("")
                                    test01 = InStr(SourceCode, "<div class=g-section sfe-break-bottom-16 overflow-floatfix>")
                                    test02 = InStr(SourceCode, "<td class=lft lm>Revenue")
                                    If test01 <> 0 And test02 <> 0 Then
                                        Dim strTFind, strTInicial, strTFinal, resT As String
                                        Dim tamTFind, tamTInicial, tamTFinal, tamTDesc As Integer
                                        Dim posTFind, posTInicial, posTFinal As Integer
                                        resT = ""
                                        strTFind = ("<div class=g-section sfe-break-bottom-16 overflow-floatfix>")
                                        strTInicial = ("<h3>")
                                        strTFinal = ("</h3>")
                                        tamTFind = Len(strTFind)
                                        tamTInicial = Len(strTInicial)
                                        tamTFinal = Len(strTFinal)
                                        posTFind = InStr(SourceCode, strTFind)
                                        posTFind = posTFind + tamTFind
                                        posTInicial = InStr(posTFind, SourceCode, strTInicial)
                                        posTInicial = posTInicial + tamTInicial
                                        posTFinal = InStr(posTInicial, SourceCode, strTFinal)
                                        tamTDesc = posTFinal - posTInicial
                                        resT = Mid(SourceCode, posTInicial, tamTDesc)
                                        txtCompany02.Text = resT
                                        lstBox.Items.Insert(0, txtMarket02.Text + " - " + symbol + " ---- " + txtCompany02.Text)
                                        lstBoxDATA.Items.Add(txtMarket02.Text + " - " + symbol + " ---- " + txtCompany02.Text)
                                        lstBoxDATA.Items.Add("")

                                        IncomeStatment(SourceCode)
                                        BalanceSheet(SourceCode)
                                        CashFlow(SourceCode)

                                        'Insert in DATABASE here

                                        txtOK02.Text = Total
                                        Total = Total + 1
                                        TotalOk = TotalOk + 1
                                        txtBoxTotalOK.Text = TotalOk
                                    Else
                                        'MsgBox("Ficheiro Não Válido")
                                        txtNOK02.Text = nok
                                        nok = nok + 1
                                        TotalNOk = TotalNOk + 1
                                        txtBoxTotalNOK.Text = TotalNOk
                                    End If

                                    '____________________________________
                                Else
                                    Do Until SkyRun = 1
                                        txtStatus02.Text = "Paused"
                                        System.Threading.Thread.Sleep(100)
                                    Loop
                                End If
                                txtStatus02.Text = "Running"
                                '____________________________________

                            Next
                        Next
                    Next
                Next
            Next
        Catch ex As Exception
            txtStatus02.Text = ("ERROR")
            btnRefresh02.Visible = True
        End Try
    End Sub
    Private Sub inThreadMainCodeExtractionNYSEAMEX()
        Try
            Dim Alfabeto() As String = {"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
            Dim Letra01, Letra02, Letra03, Letra04, Letra05 As String
            Dim Total, nok As Integer
            Dim NYSEAMEXRunOk As Integer = 0
            nok = 1
            Total = 1
            Dim SourceCode As String = ""
            txtMarket03.Text = ("NYSEAMEX")
            For Each Letra05 In Alfabeto
                For Each Letra04 In Alfabeto
                    For Each Letra03 In Alfabeto
                        For Each Letra02 In Alfabeto
                            For Each Letra01 In Alfabeto
                                '____________________________________
                                If SkyRun = 1 Then
                                    '____________________________________
                                    txtLetter05_03.Text = Letra05
                                    txtLetter04_03.Text = Letra04
                                    txtLetter03_03.Text = Letra03
                                    txtLetter02_03.Text = Letra02
                                    txtLetter01_03.Text = Letra01
                                    Dim symbol As String = (txtLetter05_03.Text) + (txtLetter04_03.Text) + (txtLetter03_03.Text) + (txtLetter02_03.Text) + (txtLetter01_03.Text)

                                    'Navegar até...
                                    'WebBrowser1.Navigate(txtURL.Text)
                                    'Sacar SourceCode
                                    'Implementar aqui verificação de erro para ligação á internet
                                    Try
                                        Dim Request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create("http://finance.google.com/finance?q=" + txtMarket03.Text + "%3A" + symbol + "&fstype=ii&hl=en")
                                        Dim Response As System.Net.HttpWebResponse = Request.GetResponse()
                                        Dim sr As System.IO.StreamReader = New System.IO.StreamReader(Response.GetResponseStream())
                                        SourceCode = sr.ReadToEnd()
                                    Catch ex As Exception
                                        NYSEAMEXRunOk = 0
                                        Do Until NYSEAMEXRunOk = 1
                                            Try
                                                txtCompany03.Text = ("Connecting...")
                                                Dim Request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create("http://finance.google.com/finance?q=" + txtMarket03.Text + "%3A" + symbol + "&fstype=ii&hl=en")
                                                Dim Response As System.Net.HttpWebResponse = Request.GetResponse()
                                                Dim sr As System.IO.StreamReader = New System.IO.StreamReader(Response.GetResponseStream())
                                                SourceCode = sr.ReadToEnd()
                                                NYSEAMEXRunOk = 1
                                            Catch ez As Exception
                                                NYSEAMEXRunOk = 0
                                                txtCompany03.Text = ("Connecting...")
                                                txtFail03.Text += 1
                                                System.Threading.Thread.Sleep(60000)
                                            End Try
                                        Loop
                                    End Try

                                    ProgressBarValue = ProgressBarValue + 1

                                    SourceCode = CharRemover(SourceCode)
                                    'GRAVAR ficheiro com fonte
                                    ' Dim FILE_NAME As String = "C:\Users\Rui\Desktop\temp.txt"
                                    ' If System.IO.File.Exists(FILE_NAME) = True Then
                                    'Dim objWriter As New System.IO.StreamWriter(FILE_NAME)
                                    ' objWriter.Write(SourceCode)
                                    'objWriter.Close()
                                    'Else
                                    ' MsgBox("File does not exist")
                                    ' End If
                                    'VERIFICAR se ficheiro ok para extrair dados
                                    Dim test01, test02 As Integer
                                    test01 = 0
                                    test02 = 0
                                    txtCompany03.Text = ("")
                                    test01 = InStr(SourceCode, "<div class=g-section sfe-break-bottom-16 overflow-floatfix>")
                                    test02 = InStr(SourceCode, "<td class=lft lm>Revenue")
                                    If test01 <> 0 And test02 <> 0 Then
                                        Dim strTFind, strTInicial, strTFinal, resT As String
                                        Dim tamTFind, tamTInicial, tamTFinal, tamTDesc As Integer
                                        Dim posTFind, posTInicial, posTFinal As Integer
                                        resT = ""
                                        strTFind = ("<div class=g-section sfe-break-bottom-16 overflow-floatfix>")
                                        strTInicial = ("<h3>")
                                        strTFinal = ("</h3>")
                                        tamTFind = Len(strTFind)
                                        tamTInicial = Len(strTInicial)
                                        tamTFinal = Len(strTFinal)
                                        posTFind = InStr(SourceCode, strTFind)
                                        posTFind = posTFind + tamTFind
                                        posTInicial = InStr(posTFind, SourceCode, strTInicial)
                                        posTInicial = posTInicial + tamTInicial
                                        posTFinal = InStr(posTInicial, SourceCode, strTFinal)
                                        tamTDesc = posTFinal - posTInicial
                                        resT = Mid(SourceCode, posTInicial, tamTDesc)
                                        txtCompany03.Text = resT
                                        lstBox.Items.Insert(0, txtMarket03.Text + " - " + symbol + " ---- " + txtCompany03.Text)
                                        lstBoxDATA.Items.Add(txtMarket03.Text + " - " + symbol + " ---- " + txtCompany03.Text)
                                        lstBoxDATA.Items.Add("")

                                        IncomeStatment(SourceCode)
                                        BalanceSheet(SourceCode)
                                        CashFlow(SourceCode)



                                        'Insert in DATABASE here

                                        txtOK03.Text = Total
                                        Total = Total + 1
                                        TotalOk = TotalOk + 1
                                        txtBoxTotalOK.Text = TotalOk
                                    Else
                                        'MsgBox("Ficheiro Não Válido")
                                        txtNOK03.Text = nok
                                        nok = nok + 1
                                        TotalNOk = TotalNOk + 1
                                        txtBoxTotalNOK.Text = TotalNOk
                                    End If
                                    '____________________________________
                                Else
                                    Do Until SkyRun = 1
                                        txtStatus03.Text = "Paused"
                                        System.Threading.Thread.Sleep(100)
                                    Loop
                                End If
                                txtStatus03.Text = "Running"
                                '____________________________________
                            Next
                        Next
                    Next
                Next
            Next
        Catch ex As Exception
            txtStatus03.Text = ("ERROR")
            btnRefresh03.Visible = True
        End Try
    End Sub
    Private Sub inThreadMainCodeExtractionTSE()
        Try
            Dim Alfabeto() As String = {"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
            Dim Letra01, Letra02, Letra03, Letra04, Letra05 As String
            Dim Total, nok As Integer
            Dim TSERunOk As Integer = 0

            'Dim BypassActive As Boolean
            'Dim Bypass As Boolean

            nok = 1
            Total = 1
            Dim SourceCode As String = ""
            txtMarket05.Text = ("TSE")
            For Each Letra05 In Alfabeto
                For Each Letra04 In Alfabeto
                    For Each Letra03 In Alfabeto
                        For Each Letra02 In Alfabeto
                            For Each Letra01 In Alfabeto


                                'Verify empty simbols
                                ' BypassActive = False
                                ' Bypass = False

                                ' If txtLetter01_04.Text = "" Then
                                'Bypass = True
                                ' ElseIf txtLetter02_04.Text = "" Then
                                'Bypass = True
                                'ElseIf txtLetter03_04.Text = "" Then
                                ' Bypass = True
                                'ElseIf txtLetter04_04.Text = "" Then
                                ' Bypass = True
                                ' ElseIf txtLetter05_04.Text = "" Then
                                ' Bypass = True
                                '  End If






                                'For Pause and Resume purpose
                                If SkyRun = 1 Then
                                    '____________________________________
                                    txtLetter05_04.Text = Letra05
                                    txtLetter04_04.Text = Letra04
                                    txtLetter03_04.Text = Letra03
                                    txtLetter02_04.Text = Letra02
                                    txtLetter01_04.Text = Letra01
                                    Dim symbol As String = (txtLetter05_04.Text) + (txtLetter04_04.Text) + (txtLetter03_04.Text) + (txtLetter02_04.Text) + (txtLetter01_04.Text)

                                    'Navegar até...
                                    'WebBrowser1.Navigate(txtURL.Text)
                                    'Sacar SourceCode
                                    'Implementar aqui verificação de erro para ligação á internet
                                    Try
                                        Dim Request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create("http://finance.google.com/finance?q=" + txtMarket05.Text + "%3A" + symbol + "&fstype=ii&hl=en")
                                        Dim Response As System.Net.HttpWebResponse = Request.GetResponse()
                                        Dim sr As System.IO.StreamReader = New System.IO.StreamReader(Response.GetResponseStream())
                                        SourceCode = sr.ReadToEnd()
                                    Catch ex As Exception
                                        TSERunOk = 0
                                        Do Until TSERunOk = 1
                                            Try
                                                txtCompany05.Text = ("Connecting...")
                                                Dim Request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create("http://finance.google.com/finance?q=" + txtMarket05.Text + "%3A" + symbol + "&fstype=ii&hl=en")
                                                Dim Response As System.Net.HttpWebResponse = Request.GetResponse()
                                                Dim sr As System.IO.StreamReader = New System.IO.StreamReader(Response.GetResponseStream())
                                                SourceCode = sr.ReadToEnd()
                                                TSERunOk = 1
                                            Catch ez As Exception
                                                TSERunOk = 0
                                                txtCompany05.Text = ("Connecting...")
                                                txtFail05.Text += 1
                                                System.Threading.Thread.Sleep(60000)
                                            End Try
                                        Loop
                                    End Try

                                    ProgressBarValue = ProgressBarValue + 1

                                    SourceCode = CharRemover(SourceCode)
                                    'GRAVAR ficheiro com fonte
                                    ' Dim FILE_NAME As String = "C:\Users\Rui\Desktop\temp.txt"
                                    ' If System.IO.File.Exists(FILE_NAME) = True Then
                                    'Dim objWriter As New System.IO.StreamWriter(FILE_NAME)
                                    ' objWriter.Write(SourceCode)
                                    'objWriter.Close()
                                    'Else
                                    ' MsgBox("File does not exist")
                                    ' End If
                                    'VERIFICAR se ficheiro ok para extrair dados
                                    Dim test01, test02 As Integer
                                    test01 = 0
                                    test02 = 0
                                    txtCompany05.Text = ("")
                                    test01 = InStr(SourceCode, "<div class=g-section sfe-break-bottom-16 overflow-floatfix>")
                                    test02 = InStr(SourceCode, "<td class=lft lm>Revenue")
                                    If test01 <> 0 And test02 <> 0 Then
                                        Dim strTFind, strTInicial, strTFinal, resT As String
                                        Dim tamTFind, tamTInicial, tamTFinal, tamTDesc As Integer
                                        Dim posTFind, posTInicial, posTFinal As Integer
                                        resT = ""
                                        strTFind = ("<div class=g-section sfe-break-bottom-16 overflow-floatfix>")
                                        strTInicial = ("<h3>")
                                        strTFinal = ("</h3>")
                                        tamTFind = Len(strTFind)
                                        tamTInicial = Len(strTInicial)
                                        tamTFinal = Len(strTFinal)
                                        posTFind = InStr(SourceCode, strTFind)
                                        posTFind = posTFind + tamTFind
                                        posTInicial = InStr(posTFind, SourceCode, strTInicial)
                                        posTInicial = posTInicial + tamTInicial
                                        posTFinal = InStr(posTInicial, SourceCode, strTFinal)
                                        tamTDesc = posTFinal - posTInicial
                                        resT = Mid(SourceCode, posTInicial, tamTDesc)
                                        txtCompany05.Text = resT
                                        lstBox.Items.Insert(0, txtMarket05.Text + " - " + symbol + " ---- " + txtCompany05.Text)
                                        lstBoxDATA.Items.Add(txtMarket05.Text + " - " + symbol + " ---- " + txtCompany05.Text)
                                        lstBoxDATA.Items.Add("")

                                        IncomeStatment(SourceCode)
                                        BalanceSheet(SourceCode)
                                        CashFlow(SourceCode)

                                        'Insert in DATABASE here

                                        txtOK05.Text = Total
                                        Total = Total + 1
                                        TotalOk = TotalOk + 1
                                        txtBoxTotalOK.Text = TotalOk
                                    Else
                                        'MsgBox("Ficheiro Não Válido")
                                        txtNOK05.Text = nok
                                        nok = nok + 1
                                        TotalNOk = TotalNOk + 1
                                        txtBoxTotalNOK.Text = TotalNOk
                                    End If
                                    '____________________________________
                                Else
                                    Do Until SkyRun = 1
                                        txtStatus05.Text = "Paused"
                                        System.Threading.Thread.Sleep(100)
                                    Loop
                                End If
                                txtStatus05.Text = "Running"
                                '____________________________________



                            Next
                        Next
                    Next
                Next
            Next
        Catch ex As Exception
            txtStatus05.Text = ("ERROR")
            btnRefresh05.Visible = True
        End Try
    End Sub
    Private Sub inThreadMainCodeExtractionPINK()
        Try
            Dim Alfabeto() As String = {"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
            Dim Letra01, Letra02, Letra03, Letra04, Letra05 As String
            Dim Total, nok As Integer
            Dim PINKRunOk As Integer = 0
            nok = 1
            Total = 1
            Dim SourceCode As String = ""
            txtMarket04.Text = ("PINK")
            For Each Letra05 In Alfabeto
                For Each Letra04 In Alfabeto
                    For Each Letra03 In Alfabeto
                        For Each Letra02 In Alfabeto
                            For Each Letra01 In Alfabeto
                                '____________________________________
                                If SkyRun = 1 Then
                                    '____________________________________
                                    txtLetter05_05.Text = Letra05
                                    txtLetter04_05.Text = Letra04
                                    txtLetter03_05.Text = Letra03
                                    txtLetter02_05.Text = Letra02
                                    txtLetter01_05.Text = Letra01
                                    Dim symbol As String = (txtLetter05_05.Text) + (txtLetter04_05.Text) + (txtLetter03_05.Text) + (txtLetter02_05.Text) + (txtLetter01_05.Text)

                                    'Navegar até...
                                    'WebBrowser1.Navigate(txtURL.Text)
                                    'Sacar SourceCode
                                    'Implementar aqui verificação de erro para ligação á internet
                                    Try
                                        Dim Request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create("http://finance.google.com/finance?q=" + txtMarket04.Text + "%3A" + symbol + "&fstype=ii&hl=en")
                                        Dim Response As System.Net.HttpWebResponse = Request.GetResponse()
                                        Dim sr As System.IO.StreamReader = New System.IO.StreamReader(Response.GetResponseStream())
                                        SourceCode = sr.ReadToEnd()
                                    Catch ex As Exception
                                        PINKRunOk = 0
                                        Do Until PINKRunOk = 1
                                            Try
                                                txtCompany04.Text = ("Connecting...")
                                                Dim Request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create("http://finance.google.com/finance?q=" + txtMarket04.Text + "%3A" + symbol + "&fstype=ii&hl=en")
                                                Dim Response As System.Net.HttpWebResponse = Request.GetResponse()
                                                Dim sr As System.IO.StreamReader = New System.IO.StreamReader(Response.GetResponseStream())
                                                SourceCode = sr.ReadToEnd()
                                                PINKRunOk = 1
                                            Catch ez As Exception
                                                PINKRunOk = 0
                                                txtCompany04.Text = ("Connecting...")
                                                txtFail04.Text += 1
                                                System.Threading.Thread.Sleep(60000)
                                            End Try
                                        Loop
                                    End Try

                                    ProgressBarValue = ProgressBarValue + 1

                                    SourceCode = CharRemover(SourceCode)
                                    'GRAVAR ficheiro com fonte
                                    ' Dim FILE_NAME As String = "C:\Users\Rui\Desktop\temp.txt"
                                    ' If System.IO.File.Exists(FILE_NAME) = True Then
                                    'Dim objWriter As New System.IO.StreamWriter(FILE_NAME)
                                    ' objWriter.Write(SourceCode)
                                    'objWriter.Close()
                                    'Else
                                    ' MsgBox("File does not exist")
                                    ' End If
                                    'VERIFICAR se ficheiro ok para extrair dados
                                    Dim test01, test02 As Integer
                                    test01 = 0
                                    test02 = 0
                                    txtCompany04.Text = ("")
                                    test01 = InStr(SourceCode, "<div class=g-section sfe-break-bottom-16 overflow-floatfix>")
                                    test02 = InStr(SourceCode, "<td class=lft lm>Revenue")
                                    If test01 <> 0 And test02 <> 0 Then
                                        Dim strTFind, strTInicial, strTFinal, resT As String
                                        Dim tamTFind, tamTInicial, tamTFinal, tamTDesc As Integer
                                        Dim posTFind, posTInicial, posTFinal As Integer
                                        resT = ""
                                        strTFind = ("<div class=g-section sfe-break-bottom-16 overflow-floatfix>")
                                        strTInicial = ("<h3>")
                                        strTFinal = ("</h3>")
                                        tamTFind = Len(strTFind)
                                        tamTInicial = Len(strTInicial)
                                        tamTFinal = Len(strTFinal)
                                        posTFind = InStr(SourceCode, strTFind)
                                        posTFind = posTFind + tamTFind
                                        posTInicial = InStr(posTFind, SourceCode, strTInicial)
                                        posTInicial = posTInicial + tamTInicial
                                        posTFinal = InStr(posTInicial, SourceCode, strTFinal)
                                        tamTDesc = posTFinal - posTInicial
                                        resT = Mid(SourceCode, posTInicial, tamTDesc)
                                        txtCompany04.Text = resT
                                        lstBox.Items.Insert(0, txtMarket04.Text + " - " + symbol + " ---- " + txtCompany04.Text)
                                        lstBoxDATA.Items.Add(txtMarket04.Text + " - " + symbol + " ---- " + txtCompany04.Text)
                                        lstBoxDATA.Items.Add("")

                                        IncomeStatment(SourceCode)
                                        BalanceSheet(SourceCode)
                                        CashFlow(SourceCode)



                                        'Insert in DATABASE here

                                        txtOK04.Text = Total
                                        Total = Total + 1
                                        TotalOk = TotalOk + 1
                                        txtBoxTotalOK.Text = TotalOk
                                    Else
                                        'MsgBox("Ficheiro Não Válido")
                                        txtNOK04.Text = nok
                                        nok = nok + 1
                                        TotalNOk = TotalNOk + 1
                                        txtBoxTotalNOK.Text = TotalNOk
                                    End If
                                    '____________________________________
                                Else
                                    Do Until SkyRun = 1
                                        txtStatus04.Text = "Paused"
                                        System.Threading.Thread.Sleep(100)
                                    Loop
                                End If
                                txtStatus04.Text = "Running"
                                '____________________________________
                            Next
                        Next
                    Next
                Next
            Next
        Catch ex As Exception
            txtStatus04.Text = ("ERROR")
            btnRefresh04.Visible = True
        End Try
    End Sub


    Private Sub inThreadCicleTotalMin()
        Dim Run As Integer = 1
        Dim last As Integer
        Dim History As String = ""
        Dim TotalMin As Integer

        Do Until Run = 0
            last = TotalOKNOk
            System.Threading.Thread.Sleep(60000)
            TotalMin = (TotalOKNOk - last)
            txtBoxTotalMin.Text = "Average Speed " & TotalMin & " /Min"
            History = TotalMin & ", " & History
            txtBoxHistory.Text = History
            If History.Length > 60 Then
                History = History.Substring(0, 60)
                txtBoxHistory.Text = History & " ..."
            End If
        Loop
    End Sub
    Private Sub inThreadCicle60Min()
        Do Until Temp = 0
            System.Threading.Thread.Sleep(3600000) '60 000 = 1 Min
            lstBoxDATA.Items.Clear()
        Loop
    End Sub

    Public Function CharRemover(ByVal SourceCode)

        'REMOVER caracter especial "
        Dim theString As String
        theString = SourceCode.Replace(Chr(34), "")
        SourceCode = theString

        Dim TempReplace As String
        TempReplace = SourceCode.Replace("<span class=chr>", "")
        SourceCode = TempReplace

        Dim TempReplaceSpan As String
        TempReplaceSpan = SourceCode.Replace("</span>", "")
        SourceCode = TempReplaceSpan

        Dim TempReplaceBld As String
        TempReplaceBld = SourceCode.Replace(" bld", "")
        SourceCode = TempReplaceBld

        Dim TempReplaceRm As String
        TempReplaceRm = SourceCode.Replace(" rm", "")
        SourceCode = TempReplaceRm

        Dim TempReplaceStylePadding As String
        TempReplaceStylePadding = SourceCode.Replace("<span style=padding-left:18px;>", "")
        SourceCode = TempReplaceStylePadding

        TotalOKNOk = TotalOk + TotalNOk
        txtBoxTotalOKNOK.Text = TotalOKNOk

        'GRAVAR ficheiro com fonte
        'Dim FILE_NAME As String = "C:\Users\Rui\Desktop\temp.txt"

        'If System.IO.File.Exists(FILE_NAME) = True Then
        'Dim objWriter As New System.IO.StreamWriter(FILE_NAME)
        'objWriter.Write(SourceCode)
        'objWriter.Close()
        'Else
        'MsgBox("File does not exist")
        ' End If

        Return (SourceCode)
    End Function

    Private Sub IncomeStatment(ByVal SourceCode)

        Dim strIncomeStatmentStart As String = ("<div id=incinterimdiv_viz class=id-incinterimdiv_viz viz_charts></div>")
        Dim strIncomeStatmentFinish As String = ("</div>")
        Dim posIncomeStatmentStart As Integer

        Dim TodasTabelasContadas As Integer = 0
        Dim TodosDadosExtraidos As Integer = 0

        Dim counter As Integer
        Dim strTitleDados As String
        Dim strAllTitles As String = ""
        Dim strValue As String = ""

        Dim strTdTitleStart As String = ("<td class=lft lm>")
        Dim strTdTitleFinish As String = ("</td>")
        Dim posTdTitleStart As Integer

        Dim tamDesconhecido As Integer

        Dim strValueStart As String = ("<td class=r>")
        Dim strValueFinish As String = ("</td>")
        Dim posValueStart As Integer

        'Encontrar IncomeStatmentStart
        posIncomeStatmentStart = InStr(SourceCode, strIncomeStatmentStart) + Len(strIncomeStatmentStart)
        lstBoxDATA.Items.Add("--- Income Statment ---")
        counter = 0
        'TodasTabelasContadas
        Do Until TodasTabelasContadas = 1
            TodosDadosExtraidos = 0
            'Encontrar tabelas e extrair informação
            'Localizar tabela seguinte
            posTdTitleStart = InStr(posIncomeStatmentStart, SourceCode, strTdTitleStart) + Len(strTdTitleStart)
            If posTdTitleStart < InStr(posIncomeStatmentStart, SourceCode, strIncomeStatmentFinish) Then
                'Retirar Informação
                posTdTitleStart = InStr(posIncomeStatmentStart, SourceCode, strTdTitleStart) + Len(strTdTitleStart)
                tamDesconhecido = InStr(posTdTitleStart, SourceCode, strTdTitleFinish) - posTdTitleStart
                strTitleDados = Mid(SourceCode, posTdTitleStart, tamDesconhecido)
                posIncomeStatmentStart = posTdTitleStart + tamDesconhecido + Len(strTdTitleFinish)
                LineToWright = strTitleDados
                Do Until TodosDadosExtraidos = 1
                    'Encontrar dados seguintes
                    posValueStart = InStr(posIncomeStatmentStart, SourceCode, strValueStart)
                    If posValueStart < InStr(posIncomeStatmentStart, SourceCode, strTdTitleFinish) Then
                        posValueStart = InStr(posIncomeStatmentStart, SourceCode, strValueStart) + Len(strValueStart)
                        tamDesconhecido = InStr(posValueStart, SourceCode, strValueFinish) - posValueStart
                        strValue = Mid(SourceCode, posValueStart, tamDesconhecido)
                        posIncomeStatmentStart = posValueStart + tamDesconhecido + Len(strValueFinish)
                        LineToWright = LineToWright + StringSeparator + strValue
                    Else
                        TodosDadosExtraidos = 1
                        lstBoxDATA.Items.Add(LineToWright)
                    End If
                Loop

            Else
                TodasTabelasContadas = 1
            End If
        Loop
    End Sub
    Private Sub BalanceSheet(ByVal SourceCode)

        Dim strIncomeStatmentStart As String = ("<div id=balinterimdiv_viz class=id-balinterimdiv_viz viz_charts></div>")
        Dim strIncomeStatmentFinish As String = ("</div>")
        Dim posIncomeStatmentStart As Integer

        Dim TodasTabelasContadas As Integer = 0
        Dim TodosDadosExtraidos As Integer = 0

        Dim counter As Integer
        Dim strTitleDados As String
        Dim strAllTitles As String = ""
        Dim strValue As String = ""

        Dim strTdTitleStart As String = ("<td class=lft lm>")
        Dim strTdTitleFinish As String = ("</td>")
        Dim posTdTitleStart As Integer

        Dim tamDesconhecido As Integer

        Dim strValueStart As String = ("<td class=r>")
        Dim strValueFinish As String = ("</td>")
        Dim posValueStart As Integer

        'Encontrar IncomeStatmentStart
        posIncomeStatmentStart = InStr(SourceCode, strIncomeStatmentStart) + Len(strIncomeStatmentStart)
        lstBoxDATA.Items.Add("")
        lstBoxDATA.Items.Add("--- Balance Sheet ---")
        counter = 0
        'TodasTabelasContadas
        Do Until TodasTabelasContadas = 1
            TodosDadosExtraidos = 0
            'Encontrar tabelas e extrair informação
            'Localizar tabela seguinte
            posTdTitleStart = InStr(posIncomeStatmentStart, SourceCode, strTdTitleStart) + Len(strTdTitleStart)
            If posTdTitleStart < InStr(posIncomeStatmentStart, SourceCode, strIncomeStatmentFinish) Then
                'Retirar Informação
                posTdTitleStart = InStr(posIncomeStatmentStart, SourceCode, strTdTitleStart) + Len(strTdTitleStart)
                tamDesconhecido = InStr(posTdTitleStart, SourceCode, strTdTitleFinish) - posTdTitleStart
                strTitleDados = Mid(SourceCode, posTdTitleStart, tamDesconhecido)
                posIncomeStatmentStart = posTdTitleStart + tamDesconhecido + Len(strTdTitleFinish)
                LineToWright = strTitleDados
                Do Until TodosDadosExtraidos = 1
                    'Encontrar dados seguintes
                    posValueStart = InStr(posIncomeStatmentStart, SourceCode, strValueStart)
                    If posValueStart < InStr(posIncomeStatmentStart, SourceCode, strTdTitleFinish) Then
                        posValueStart = InStr(posIncomeStatmentStart, SourceCode, strValueStart) + Len(strValueStart)
                        tamDesconhecido = InStr(posValueStart, SourceCode, strValueFinish) - posValueStart
                        strValue = Mid(SourceCode, posValueStart, tamDesconhecido)
                        posIncomeStatmentStart = posValueStart + tamDesconhecido + Len(strValueFinish)
                        LineToWright = LineToWright + StringSeparator + strValue
                    Else
                        TodosDadosExtraidos = 1
                        lstBoxDATA.Items.Add(LineToWright)
                    End If
                Loop

            Else
                TodasTabelasContadas = 1
            End If
        Loop
    End Sub
    Private Sub CashFlow(ByVal SourceCode)

        Dim strIncomeStatmentStart As String = ("<div id=casinterimdiv_viz class=id-casinterimdiv_viz viz_charts></div>")
        Dim strIncomeStatmentFinish As String = ("</div>")
        Dim posIncomeStatmentStart As Integer

        Dim TodasTabelasContadas As Integer = 0
        Dim TodosDadosExtraidos As Integer = 0

        Dim counter As Integer
        Dim strTitleDados As String
        Dim strAllTitles As String = ""
        Dim strValue As String = ""

        Dim strTdTitleStart As String = ("<td class=lft lm>")
        Dim strTdTitleFinish As String = ("</td>")
        Dim posTdTitleStart As Integer

        Dim tamDesconhecido As Integer

        Dim strValueStart As String = ("<td class=r>")
        Dim strValueFinish As String = ("</td>")
        Dim posValueStart As Integer

        'Encontrar IncomeStatmentStart
        posIncomeStatmentStart = InStr(SourceCode, strIncomeStatmentStart) + Len(strIncomeStatmentStart)
        lstBoxDATA.Items.Add("")
        lstBoxDATA.Items.Add("--- Cash Flow ---")

        counter = 0
        'TodasTabelasContadas
        Do Until TodasTabelasContadas = 1
            TodosDadosExtraidos = 0
            'Encontrar tabelas e extrair informação
            'Localizar tabela seguinte
            posTdTitleStart = InStr(posIncomeStatmentStart, SourceCode, strTdTitleStart) + Len(strTdTitleStart)
            If posTdTitleStart < InStr(posIncomeStatmentStart, SourceCode, strIncomeStatmentFinish) Then
                'Retirar Informação
                posTdTitleStart = InStr(posIncomeStatmentStart, SourceCode, strTdTitleStart) + Len(strTdTitleStart)
                tamDesconhecido = InStr(posTdTitleStart, SourceCode, strTdTitleFinish) - posTdTitleStart
                strTitleDados = Mid(SourceCode, posTdTitleStart, tamDesconhecido)
                posIncomeStatmentStart = posTdTitleStart + tamDesconhecido + Len(strTdTitleFinish)
                LineToWright = strTitleDados
                Do Until TodosDadosExtraidos = 1
                    'Encontrar dados seguintes
                    posValueStart = InStr(posIncomeStatmentStart, SourceCode, strValueStart)
                    If posValueStart < InStr(posIncomeStatmentStart, SourceCode, strTdTitleFinish) Then
                        posValueStart = InStr(posIncomeStatmentStart, SourceCode, strValueStart) + Len(strValueStart)
                        tamDesconhecido = InStr(posValueStart, SourceCode, strValueFinish) - posValueStart
                        strValue = Mid(SourceCode, posValueStart, tamDesconhecido)
                        posIncomeStatmentStart = posValueStart + tamDesconhecido + Len(strValueFinish)
                        LineToWright = LineToWright + StringSeparator + strValue
                    Else
                        TodosDadosExtraidos = 1
                        lstBoxDATA.Items.Add(LineToWright)
                    End If
                Loop
            Else
                TodasTabelasContadas = 1
                lstBoxDATA.Items.Add("___________________________________________________________________________")
                lstBoxDATA.Items.Add("")
                lstBoxDATA.Items.Add("")
            End If
        Loop
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    Private Sub btnStartNYSE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtStatus01.Text = "Running"
        threadNYSE = New System.Threading.Thread(AddressOf inThreadMainCodeExtractionNYSE)
        threadNYSE.Start()
    End Sub
    Private Sub btnStartNASDAQ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtStatus02.Text = "Running"
        threadNASDAQ = New System.Threading.Thread(AddressOf inThreadMainCodeExtractionNASDAQ)
        threadNASDAQ.Start()
    End Sub
    Private Sub btnStartNYSEAMEX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtStatus03.Text = "Running"
        threadNYSEAMEX = New System.Threading.Thread(AddressOf inThreadMainCodeExtractionNYSEAMEX)
        threadNYSEAMEX.Start()
    End Sub
    Private Sub btnStartPINK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtStatus04.Text = "Running"
        threadPINK = New System.Threading.Thread(AddressOf inThreadMainCodeExtractionPINK)
        threadPINK.Start()
    End Sub
    Private Sub btnStartTSE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtStatus05.Text = "Running"
        threadTSE = New System.Threading.Thread(AddressOf inThreadMainCodeExtractionTSE)
        threadTSE.Start()
    End Sub
    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        End
    End Sub
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        txtFail01.Text = 0
        txtFail02.Text = 0
        txtFail03.Text = 0
        txtFail04.Text = 0
        txtFail05.Text = 0
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub btnPauseSky_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStartPauseResume.Click

        If txtLetter01_01.Text = "" And txtLetter02_01.Text = "" And txtLetter03_01.Text = "" And txtLetter04_01.Text = "" And txtLetter05_01.Text = "" Then

            btnRefresh01.Visible = False
            btnRefresh02.Visible = False
            btnRefresh03.Visible = False
            btnRefresh04.Visible = False
            btnRefresh05.Visible = False

            txtStatus01.Text = "Running"
            threadNYSE = New System.Threading.Thread(AddressOf inThreadMainCodeExtractionNYSE)
            threadNYSE.Start()
            txtStatus02.Text = "Running"
            threadNASDAQ = New System.Threading.Thread(AddressOf inThreadMainCodeExtractionNASDAQ)
            threadNASDAQ.Start()
            txtStatus03.Text = "Running"
            threadNYSEAMEX = New System.Threading.Thread(AddressOf inThreadMainCodeExtractionNYSEAMEX)
            threadNYSEAMEX.Start()
            txtStatus04.Text = "Running"
            threadPINK = New System.Threading.Thread(AddressOf inThreadMainCodeExtractionPINK)
            threadPINK.Start()
            txtStatus05.Text = "Running"
            threadTSE = New System.Threading.Thread(AddressOf inThreadMainCodeExtractionTSE)
            threadTSE.Start()

            'threadCycleRun = New System.Threading.Thread(AddressOf ProgressBar)
            'threadCycleRun.Start()

            threadTotalMin = New System.Threading.Thread(AddressOf inThreadCicleTotalMin)
            threadTotalMin.Start()
            threadCicle60Min = New System.Threading.Thread(AddressOf inThreadCicle60Min)
            threadCicle60Min.Start()

        ElseIf SkyRun = 0 Then
            SkyRun = 1
        Else
            SkyRun = 0
        End If

    End Sub

    Private Sub lstBoxDATA_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstBoxDATA.SelectedIndexChanged

    End Sub
End Class



