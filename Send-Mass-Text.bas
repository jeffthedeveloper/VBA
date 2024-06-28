Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub enviarFormularioEmMassa()

    'Seleciona sheet daods
    Sheets("Dados").Select

    Dim driveChorme As WebDriver
    
    
    
    Dim Cell As Range
    col = 2
    
    Dim ultimaLinha As String
    ultimaLinha = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    'Coloca a palavra Parar na ultima linha
    Range("A" & ultimaLinha) = "Parar"
    
    On Error GoTo limpa
    
    'Faça Até encontrar a palavra Parar
    Do Until Sheets("Dados").Cells(col, 1) = "Parar"
    
    
            Set driveChorme = New ChromeDriver
            
            driveChorme.Get "https://pt.surveymonkey.com/r/Y9Y6FFR"
            
            'Aguarda 2 segundos para dar tempo do site abrir
            Application.Wait (Now + TimeValue("00:00:02"))
            
            'Seleciona o campo Name do formulário
            Set selecionaCampoNome = driveChorme.FindElementByName("683928983")
                    
            'Envia os dados para o campo name
            selecionaCampoNome.SendKeys (Sheets("Dados").Cells(col, 1))
            
            '-------------------------------------------------------------
            
            'Aguarda 2 segundos para dar tempo do site abrir
            Application.Wait (Now + TimeValue("00:00:02"))
            
            'Seleciona o campo Email do formulário
            Set selecionaCampoEmail = driveChorme.FindElementByName("683932318")
            
            'Envia os dados para o campo email
            selecionaCampoEmail.SendKeys (Sheets("Dados").Cells(col, 2))
            
            
            '-------------------------------------------------------------
            
            'Aguarda 2 segundos para dar tempo do site abrir
            Application.Wait (Now + TimeValue("00:00:02"))
            
            'Seleciona o campo Telefone do formulário
            Set selecionaCampoTelefone = driveChorme.FindElementByName("683930688")
            
            'Envia os dados para o campo email
            selecionaCampoTelefone.SendKeys (Sheets("Dados").Cells(col, 3))
            
            '-------------------------------------------------------------
            
            'Aguarda 2 segundos para dar tempo do site abrir
            Application.Wait (Now + TimeValue("00:00:02"))
            
            'Seleciona o campo Sobre do formulário
            Set selecionaCampoSobre = driveChorme.FindElementByName("683932969")
            
            'Envia os dados para o campo sobre
            selecionaCampoSobre.SendKeys (Sheets("Dados").Cells(col, 5))
            
            
            
            'Aguarda 2 segundos para dar tempo do site abrir
            Application.Wait (Now + TimeValue("00:00:02"))
            
            If Sheets("Dados").Cells(col, 4) = "Masculino" Then
            
                'Preenche Radio Button Masculino
                Set radioButtonMasculino = driveChorme.FindElementById("683931881_4497366118_label")
                radioButtonMasculino.Click
            
            Else
            
                'Preenche Radio Button Feminino
                Set radioButtonFeminino = driveChorme.FindElementById("683931881_4497366119_label")
                radioButtonFeminino.Click
            
            End If
            
            'Aguarda 2 segundos para dar tempo do site abrir
            Application.Wait (Now + TimeValue("00:00:02"))
            
            col = col + 1
            
            'Enviar informações formulario clicando no botão Enviar Dados
            Set enviarInformacoes = driveChorme.FindElementByXPath("//*[@id=""patas""]/main/article/section/form/div[2]/button")
            enviarInformacoes.Click
            
            'Aguarda 3 segundos para dar tempo do site abrir
            Application.Wait (Now + TimeValue("00:00:03"))
            
            'Fecha o site do formulario depois que envia
            driveChorme.Close
            
    'Clico = Repita
    Loop
    
    MsgBox "Formularios preenchidos com sucesso!!!"
            
        
        
        
            
            
limpa:
            
    
    

End Sub

