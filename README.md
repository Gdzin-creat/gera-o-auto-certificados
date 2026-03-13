# 🎓 Automatic Certificate Generator (Excel + VBA + PowerPoint)

Automação para geração de certificados em **PDF** a partir de uma lista de participantes em **Excel**, utilizando um **modelo de certificado no PowerPoint**.

Este projeto foi desenvolvido para resolver um problema real durante a organização de um **evento da LUEB**, onde era necessário emitir dezenas de certificados rapidamente, mantendo padronização e evitando erros manuais.

---

# 🚀 Overview

O script utiliza **VBA dentro do Excel** para automatizar todo o processo de geração de certificados.

Ele executa as seguintes tarefas:

1. Lê uma lista de participantes em uma planilha Excel.
2. Abre um modelo de certificado no PowerPoint.
3. Substitui automaticamente o placeholder `<NOME>` pelo nome do participante.
4. Exporta cada certificado como **PDF individual**.
5. Salva todos os certificados em uma pasta específica.

O resultado é a geração automática de dezenas (ou centenas) de certificados em poucos segundos.

---

# 📸 Workflow

```text
Excel (lista de participantes)
        │
        ▼
VBA Script
        │
        ▼
PowerPoint Template
        │
        ▼
Substituição do placeholder <NOME>
        │
        ▼
Exportação automática
        │
        ▼
PDF individual para cada participante
```

---

# 📂 Project Structure

```text
certificados/
│
├── modelo_certificado.pptx
│
├── pdf_certificados/
│
└── script_excel.xlsm
```

* **modelo_certificado.pptx** → template do certificado
* **pdf_certificados/** → pasta onde os PDFs gerados serão salvos
* **script_excel.xlsm** → planilha contendo o script VBA e a lista de participantes

---

# ⚙️ Requirements

Para executar o script é necessário:

* Microsoft **Excel**
* Microsoft **PowerPoint**
* Macros habilitadas no Excel
* Arquivo Excel salvo como **.xlsm**

---

# 📋 Excel Data Format

A planilha deve seguir o seguinte formato:

| ID | Nome           |
| -- | -------------- |
| 1  | João Silva     |
| 2  | Maria Santos   |
| 3  | Pedro Oliveira |

Regras importantes:

* A **linha 1 deve ser o cabeçalho**
* Os **nomes devem estar na coluna B**
* O script percorre da **linha 2 até a linha 76** (pode ser alterado no código)

---

# 🎨 PowerPoint Template

O modelo de certificado deve conter o placeholder:

```text
<NOME>
```

Exemplo:

```
Certificamos que <NOME> participou do evento...
```

Durante a execução do script, `<NOME>` será substituído automaticamente.

---

# 📄 Output

O script gera automaticamente:

```text
Joao Silva.pdf
Maria Santos.pdf
Pedro Oliveira.pdf
```

Todos os arquivos são salvos na pasta:

```
pdf_certificados/
```

---

# 🧠 Key Features

* ✅ Geração automática de certificados
* ✅ Exportação direta para PDF
* ✅ Prevenção de caracteres inválidos em nomes de arquivos
* ✅ Preservação do template original
* ✅ Processamento em lote de participantes

---

# 🔧 Customization

Algumas partes do script podem ser facilmente modificadas:

### Alterar quantidade de participantes

```vba
For i = 2 To 76
```

### Alterar caminho do modelo

```vba
caminhoModelo = "caminho/do/modelo.pptx"
```

### Alterar pasta de saída

```vba
pastaSalvar = "caminho/para/salvar/pdf/"
```

---

# 💡 Possible Improvements

Este projeto pode evoluir para funcionalidades mais avançadas:

* envio automático de certificados por e-mail
* interface gráfica para configuração
* leitura automática da quantidade de participantes
* integração com Google Forms ou inscrições online
* geração de QR Code nos certificados

---

# 🎯 Motivation

Durante a organização de um **evento acadêmico da LUEB**, surgiu a necessidade de emitir rapidamente dezenas de certificados personalizados.

Gerar manualmente cada certificado no PowerPoint seria extremamente demorado.
Este script foi criado para automatizar completamente esse processo.

O resultado foi uma **redução drástica no tempo de emissão dos certificados**.

---

# 📜 License

Este projeto está disponível para uso educacional, acadêmico e para automações administrativas simples.

Sinta-se livre para adaptar e melhorar a ferramenta.

---
# 👨‍💻 Author

Gdzin-creat
