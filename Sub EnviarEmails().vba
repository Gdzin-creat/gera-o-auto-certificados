Option Explicit

Sub EnviarEmails()

    Dim nome As String
    Dim email As String
    Dim caminhoPDF As String
    Dim arquivoPDF As String
    
    Dim i As Long
    Dim ultimaLinha As Long
    
    Dim objEmail As Object
    Dim objConf As Object

    On Error GoTo ErroGeral

    ' Caminho da pasta dos certificados
    caminhoPDF = "C:\Users\Godoizin\OneDrive\Área de Trabalho\certificados\pdf_certificados\"

    ultimaLinha = Cells(Rows.Count, 2).End(xlUp).Row

    For i = 2 To ultimaLinha

        Application.StatusBar = "Enviando " & (i - 1) & " de " & (ultimaLinha - 1)

        nome = Trim(Cells(i, 2).Value)
        email = Trim(Cells(i, 3).Value)

        If nome <> "" And email <> "" Then

            On Error GoTo ErroEnvio

            ' Monta o nome do arquivo PDF
            arquivoPDF = caminhoPDF & nome & ".pdf"

            ' Verifica se o arquivo existe
            If Dir(arquivoPDF) = "" Then
                Cells(i, 4).Value = "PDF NÃO ENCONTRADO"
                GoTo Proximo
            End If

            Set objEmail = CreateObject("CDO.Message")
            Set objConf = CreateObject("CDO.Configuration")

            objConf.Load -1

            With objConf.Fields
                .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
                .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
                .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
                .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
                .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "luebmarketing@gmail.com"
                .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "SUA_SENHA_DE_APLICATIVO"
                .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
                .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
                .Update
            End With

            With objEmail
                Set .Configuration = objConf
                
                .To = email
                .From = "luebmarketing@gmail.com"
                .Subject = "Certificado de Participação – Evento Holambra | CEIC FCA"
                
                .HTMLBody = "<p>Prezado(a) <b>" & nome & "</b>,</p>" & _
                            "<p>É com satisfação que a Liga Universitária de Empreendedorismo de Botucatu (LUEB) encaminha o seu certificado de participação referente ao evento realizado em Holambra, no Centro de Exposição, Inovação e Cultura (CEIC) da FCA.</p>" & _
                            "<p>O seu certificado encontra-se anexado a este e-mail.</p>" & _
                            "<br>" & _
                            "<p><b>Desafio Holambra:</b><br>" & _
                            "Informamos que estão abertas as inscrições para o Desafio Holambra. Convidamos você a participar.<br>" & _
                            "Acesse: <a href='https://pt.surveymonkey.com/r/Y7HXCQJ'>Clique aqui para se inscrever</a></p>" & _
                            "<br>" & _
                            "<p>Em caso de dúvidas, permanecemos à disposição.</p>" & _
                            "<br><p>Atenciosamente,<br>LUEB</p>"

                ' 🔥 ANEXO DO PDF
                .AddAttachment arquivoPDF

                .Send
            End With

            Cells(i, 4).Value = "ENVIADO"

Limpeza:
            Set objEmail = Nothing
            Set objConf = Nothing

            On Error GoTo 0

            Application.Wait Now + TimeValue("00:00:12")

        End If

        DoEvents
        GoTo Proximo

ErroEnvio:
        Cells(i, 4).Value = "ERRO ENVIO"
        Resume Limpeza

Proximo:

    Next i

    Application.StatusBar = False
    MsgBox "Envio finalizado!"

    Exit Sub

ErroGeral:
    MsgBox "Erro geral: " & Err.Description

End Sub