Sub GerarCertificados()

Dim pptApp As Object
Dim pptPres As Object
Dim slide As Object
Dim shp As Object

Dim nome As String
Dim nomeArquivo As String

Dim caminhoModelo As String
Dim pastaSalvar As String

Dim i As Long

caminhoModelo = "C:\Users\Godoizin\OneDrive\Área de Trabalho\certificados\modelo_certificado.pptx"
pastaSalvar = "C:\Users\Godoizin\OneDrive\Área de Trabalho\certificados\pdf_certificados\"

If Dir(pastaSalvar, vbDirectory) = "" Then MkDir pastaSalvar

'abrir PowerPoint
On Error Resume Next
Set pptApp = GetObject(, "PowerPoint.Application")
On Error GoTo 0

If pptApp Is Nothing Then
Set pptApp = CreateObject("PowerPoint.Application")
End If

pptApp.Visible = True

'LOOP pelos nomes usar de acordo com a quantidade de linhas que você precisa, no exemplo tem 75 linhas, então começa do 2 até 76 (considerando que a primeira linha é o cabeçalho)
For i = 2 To 76

nome = Trim(Cells(i, 2).Value)

If nome <> "" Then

nome = Application.WorksheetFunction.Proper(nome)

'nome do arquivo
nomeArquivo = nome

'limpar caracteres inválidos
nomeArquivo = Replace(nomeArquivo, "/", "")
nomeArquivo = Replace(nomeArquivo, "\", "")
nomeArquivo = Replace(nomeArquivo, ":", "")
nomeArquivo = Replace(nomeArquivo, "*", "")
nomeArquivo = Replace(nomeArquivo, "?", "")
nomeArquivo = Replace(nomeArquivo, """", "")
nomeArquivo = Replace(nomeArquivo, "<", "")
nomeArquivo = Replace(nomeArquivo, ">", "")
nomeArquivo = Replace(nomeArquivo, "|", "")

'ABRE UMA NOVA CÓPIA DO MODELO E INSERE O NOME PARA NÃO ALTERAR O MODELO ORIGINAL
Set pptPres = pptApp.Presentations.Open(caminhoModelo, , , False)

Set slide = pptPres.Slides(1)

'procura o placeholder
For Each shp In slide.Shapes

If shp.HasTextFrame Then

If InStr(shp.TextFrame.TextRange.Text, "<NOME>") > 0 Then

shp.TextFrame.TextRange.Text = Replace(shp.TextFrame.TextRange.Text, "<NOME>", nome)

shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2

End If

End If

Next shp

'exporta PDF
pptPres.SaveAs pastaSalvar & nomeArquivo & ".pdf", 32

'fecha SEM salvar alterações no modelo
pptPres.Close

End If

Next i

pptApp.Quit

MsgBox "Todos certificados foram gerados!"

End Sub
