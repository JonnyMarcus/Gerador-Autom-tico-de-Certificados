# 🎓 Certifica

**Transforme uma planilha de participantes em certificados prontos para compartilhar.**

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Interface-Streamlit-FF4B4B?logo=streamlit&logoColor=white)](https://streamlit.io/)
[![Deploy](https://img.shields.io/badge/Deploy-Streamlit%20Community%20Cloud-7A5AF8)](https://share.streamlit.io/)

> Uma ferramenta simples para personalizar, gerar e distribuir certificados em lote.

**[Abrir o Certifica](https://arpalq47szetr6bwvpxtzy.streamlit.app/)** · **[Ver o código](https://github.com/JonnyMarcus/Gerador-Autom-tico-de-Certificados)**

## Prévia da interface

![Prévia ilustrativa da interface do Certifica, com etapas, envio de arquivos e personalização do certificado](docs/preview-certifica.svg)

*Prévia ilustrativa da interface. A aplicação ao vivo pode variar conforme os arquivos enviados e as opções selecionadas.*

## O que você pode fazer

- Gerar certificados individuais a partir de planilhas Excel.
- Usar modelos PDF ou PowerPoint.
- Ajustar posição, cor, tamanho e estilo do nome com uma prévia interativa.
- Incluir campos adicionais, logo, assinatura e QR Code verificável.
- Baixar certificados em lote ou enviar cada certificado por e-mail.
- Usar a interface web ou a linha de comando.

## Como usar a interface

1. Envie uma planilha `.xlsx` com uma coluna `nome`.
2. Envie um modelo PDF ou PowerPoint.
3. Personalize a aparência e confira a prévia.
4. Gere e baixe o lote. O envio por e-mail é opcional e pede confirmação antes de enviar.

Para começar sem arquivos próprios, use os botões de download dos modelos de planilha e certificado disponíveis no app.

## Executar localmente

Requisitos: Python 3.10 ou superior.

```bash
git clone https://github.com/JonnyMarcus/Gerador-Autom-tico-de-Certificados.git
cd Gerador-Autom-tico-de-Certificados
python -m venv .venv
```

Ative o ambiente virtual e instale as dependências da interface:

```powershell
# Windows PowerShell
.venv\Scripts\Activate.ps1
```

```bash
# Linux ou macOS
source .venv/bin/activate
```

```bash
python -m pip install -e ".[web]"
streamlit run app.py
```

## Linha de comando

Instale o pacote e gere arquivos PowerPoint individuais a partir de uma planilha e de um modelo `.pptx`:

```bash
python -m pip install -e .
gerar-certificados \
  --planilha data/participantes.xlsx \
  --modelo assets/certificado_base.pptx \
  --saida output \
  --coluna-nome nome \
  --campos curso carga_horaria
```

Os marcadores do PowerPoint devem corresponder aos cabeçalhos da planilha, por exemplo `{nome}`, `{curso}` e `{carga_horaria}`. Mantenha cada marcador como texto contínuo dentro de um parágrafo para preservar melhor a formatação.

Consulte todas as opções com `gerar-certificados --help`.

## Envio por e-mail com Gmail

O envio é opcional e usa SMTP. Para Gmail, crie uma senha de app na Conta Google com a verificação em duas etapas ativa. No Streamlit Community Cloud, abra **Settings → Secrets** do app e configure:

```toml
smtp_host = "smtp.gmail.com"
smtp_port = 587
smtp_username = "seu-email@gmail.com"
smtp_password = "sua-senha-de-app"
email_from = "seu-email@gmail.com"
smtp_use_ssl = false
smtp_starttls = true
```

Nunca coloque a senha de app neste README, no código ou no GitHub. O app valida e mostra os destinatários antes do envio; é necessário confirmar para enviar. Sem SMTP configurado, a geração e o download continuam disponíveis.

Para outro servidor SMTP, os mesmos campos podem ser configurados com os dados do provedor. Localmente, crie `.streamlit/secrets.toml` usando [`.streamlit/secrets.toml.example`](.streamlit/secrets.toml.example); esse arquivo com credenciais é ignorado pelo Git.

## QR Codes de verificação

Para assinar QR Codes, configure `qr_secret` nos Secrets do Streamlit com pelo menos 32 caracteres aleatórios e `public_url` com o endereço público do app. Gere uma chave com:

```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

Mantenha a mesma chave para validar certificados já emitidos: trocar o segredo invalida as assinaturas anteriores. O QR pode conter o nome, a data e os campos escolhidos para verificação.

## Deploy

O app está preparado para o [Streamlit Community Cloud](https://share.streamlit.io/). Conecte o repositório GitHub, selecione a branch `main` e informe `app.py` como arquivo principal. Guarde as credenciais SMTP e a chave do QR na aba **Secrets** do app; nunca as versione.

## Limites e privacidade

- Planilha: até 5 MB.
- Modelo PDF: até 10 MB.
- Lote: até 500 participantes.
- Saída estimada: até 100 MB.
- O histórico de downloads dura somente enquanto a sessão do navegador estiver aberta.
- Os arquivos são processados no servidor em diretórios temporários. Evite enviar documentos confidenciais para uma instância pública.

## Testes

```bash
python -m pip install -e ".[web]"
python -m unittest discover -s tests -v
```

## Estrutura do projeto

```text
app.py                         Aplicação web Streamlit
src/gerador_certificados/      Geração, validação e envio
  cli.py                       Interface de linha de comando
  emailer.py                   Envio individual via SMTP
  pdf_certificado.py           Personalização de PDFs
  certificado.py               Preenchimento de modelos PowerPoint
  seguranca_qr.py              Assinatura e validação de QR Codes
  editor_posicao/              Editor visual de posição e fonte manuscrita
tests/                         Testes automatizados
assets/                        Recursos e modelos do projeto
```

---

Feito com cuidado por **[Jonny Marcus](https://github.com/JonnyMarcus)**.
