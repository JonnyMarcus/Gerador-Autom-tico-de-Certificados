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

A interface permite enviar uma planilha `.xlsx` e um certificado base `.pdf`, conferir uma amostra dos participantes, posicionar o nome arrastando-o sobre a primeira página e visualizar a prévia antes de gerar. A coluna dos nomes deve se chamar `nome`; outras colunas (por exemplo `curso`, `carga_horaria` e `data`) podem ser selecionadas e posicionadas individualmente. A tela avisa sobre nomes repetidos, campos obrigatórios vazios e texto detectável perto das áreas de impressão. Também oferece modelos de planilha e certificado para baixar.

A tipografia oferece famílias modernas (Helvetica), clássicas (Times) e monoespaçadas (Courier), cada uma com variações regular, negrito, itálico e negrito itálico, além da fonte manuscrita [Great Vibes](https://github.com/google/fonts/tree/main/ofl/greatvibes). Great Vibes é distribuída sob a SIL Open Font License 1.1; a fonte e a licença acompanham os arquivos do editor em `src/gerador_certificados/editor_posicao/`. A animação de entrada usa GSAP e respeita a configuração de movimento reduzido do dispositivo. A primeira página do PDF é renderizada para o editor visual pelo PyMuPDF. Os arquivos enviados são processados em diretório temporário durante a sessão.

A interface limita a planilha a 5 MB, o certificado a 10 MB, a geração a 500 participantes e o tamanho estimado da saída a 100 MB.

Para executar localmente:

```bash
python -m pip install -e ".[web]"
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
Web: planilha XLSX + certificado PDF → validação → nome na 1ª página → ZIP de PDFs
CLI: planilha XLSX + modelo PPTX → validação → preenchimento de marcadores → PPTX individuais
```

## Próximas etapas

1. Criar arquivos de exemplo válidos; os arquivos de modelo atualmente no repositório estão vazios.
2. Permitir ajustar a posição do nome arrastando-o diretamente sobre a prévia.

## Testes

```bash
python -m pip install -e ".[web]"
python -m unittest discover -s tests -v
```
