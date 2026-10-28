# Gerador automático de certificados

Aplicação com interface web e linha de comando para gerar certificados personalizados a partir de uma planilha Excel.

## Estrutura

```text
assets/                         Modelos e recursos visuais
data/                           Planilhas de entrada e configuração
output/                         Certificados gerados (não versionados)
app.py                          Interface web Streamlit
src/gerador_certificados/
  cli.py                        Interface de linha de comando
  config.py                     Caminhos e configurações padrão
  planilha.py                   Leitura do Excel
  validacao.py                  Validação dos dados de entrada
  certificado.py                Substituição dos marcadores no PowerPoint
  pdf_certificado.py            Aplicação dos nomes sobre o PDF
  exportador.py                 Criação e nomeação dos arquivos
  logging_config.py             Configuração de logs
```

## Aplicação web

A interface permite enviar uma planilha `.xlsx` e um certificado base `.pdf`, conferir uma amostra dos participantes e baixar os certificados personalizados em um arquivo ZIP. A coluna dos nomes é fixa e deve se chamar `nome`. O nome é aplicado ao centro da primeira página do PDF; deixe essa área livre no modelo. A tela usa uma animação de entrada GSAP, respeitando a configuração de movimento reduzido do dispositivo. Os arquivos enviados são processados em diretório temporário durante a sessão.

Para executar localmente:

```bash
python -m pip install -e .
streamlit run app.py
```

O deploy recomendado é o Streamlit Community Cloud, que oferece hospedagem gratuita conectada a um repositório GitHub. Para publicar, conecte o GitHub em [share.streamlit.io](https://share.streamlit.io), escolha este repositório, a branch e `app.py` como arquivo principal. A documentação está em [Deploy your app](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy).

O app processa os arquivos no servidor e usa diretórios temporários. Ele não deve ser usado para compartilhar documentos confidenciais em uma instância pública.

## Requisitos

- Python 3.10 ou superior
- Um arquivo `.xlsx` com cabeçalho na primeira linha e uma coluna `nome`
- Para a interface web: um modelo `.pdf` com espaço livre ao centro e uma planilha com a coluna `nome`
- Para a CLI: um modelo `.pptx` com marcadores entre chaves, como `{nome}`, `{curso}` e `{carga_horaria}`

Na CLI, os nomes dos marcadores devem corresponder aos cabeçalhos da planilha, sem distinção entre maiúsculas e minúsculas. Para preservar a formatação, deixe cada marcador como texto contínuo no mesmo parágrafo do PowerPoint.

## Instalação

```bash
python -m venv .venv
# Windows PowerShell
.venv\Scripts\Activate.ps1
# Linux/macOS: source .venv/bin/activate
python -m pip install -e .
```

## Uso

Coloque a planilha em `data/participantes.xlsx` e o modelo em `assets/certificado_base.pptx`, depois execute:

```bash
gerar-certificados
```

Para informar caminhos e validar colunas adicionais:

```bash
gerar-certificados --planilha data/participantes.xlsx --modelo assets/certificado_base.pptx --saida output --coluna-nome nome --campos curso carga_horaria
```

Ajuda do comando:

```bash
gerar-certificados --help
```

Na interface web, cada certificado é gerado em `.pdf` e entregue em um `.zip`. A CLI continua gerando arquivos `.pptx` a partir de modelos PowerPoint.

## Exemplo de planilha

| nome | curso | carga_horaria |
|---|---|---|
| Ana Souza | Primeiros Socorros | 8 horas |
| João Lima | Primeiros Socorros | 8 horas |

## Fluxo da aplicação

```text
Interface web ou CLI → leitura da planilha → validação → preenchimento do modelo → exportação dos PPTX
```

## Próximas etapas

1. Criar um modelo PowerPoint com o layout final e os marcadores.
2. Preencher a planilha Excel com os participantes reais.
3. Definir a exportação em PDF, caso necessária (por exemplo, via LibreOffice).
4. Adicionar testes automatizados para planilha, validação e substituição dos marcadores.
5. Se necessário, incluir exportação em PDF e um limite configurável para a quantidade de participantes.
