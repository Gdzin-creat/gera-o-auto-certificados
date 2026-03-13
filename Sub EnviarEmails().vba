Sub EnviarEmails()

Dim nome As String
Dim email As String

Dim i As Long
Dim ultimaLinha As Long

Dim objEmail As Object
Dim objConf As Object

ultimaLinha = Cells(Rows.Count, 2).End(xlUp).Row

For i = 2 To ultimaLinha

Application.StatusBar = "Enviando " & (i - 1) & " de " & (ultimaLinha - 1)

nome = Trim(Cells(i, 2).Value)
email = Trim(Cells(i, 3).Value)

If nome <> "" And email <> "" Then

On Error Resume Next

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
    .Subject = "Mensagem da LUEB"
    
    .HTMLBody = "<p>Olá <b>" & nome & "</b>,</p>" & _
                "<p>Este é um teste de envio automático de emails pela Liga Universitária de Empreendedorismo de Botucatu (LUEB).</p>" & _
                "<p>Se você recebeu este email, o sistema está funcionando corretamente.</p>" & _
                "<br><p>Atenciosamente,<br>LUEB</p>"

    .Send

End With

If Err.Number <> 0 Then
    Cells(i, 4).Value = "ERRO ENVIO"
    Err.Clear
Else
    Cells(i, 4).Value = "ENVIADO"
End If

Application.Wait Now + TimeValue("00:00:12")

End If

DoEvents

Next i

Application.StatusBar = False

MsgBox "Envio finalizado!"

End Sub