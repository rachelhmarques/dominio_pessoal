# Domínio Pessoal

Aplicação em Python para transformar o relatório `Resumo Mensal` do setor pessoal em lançamentos do `DOMINIO LANCAMENTO`, com suporte a consolidação opcional do arquivo complementar `Resumo Mensal13`, além de:

- prévia dos lançamentos na tela
- memória de cálculo por competência
- mapeamento editável
- exportação em CSV
- exportação em Excel
- interface pronta para deploy no Streamlit Cloud

## Arquivos principais

- `streamlit_app.py`: interface principal para uso local e deploy no Streamlit
- `dominio_parser.py`: regras de leitura, cálculo, mapeamento e exportação
- `verify_sample.py`: validação do caso de referência da filial 41

## Requisitos

- Python 3.12 ou compatível
- Dependências de `requirements.txt`

Observações:

- O parser lê `.xls` em Python puro.
- Para `.xlsx`, também existe leitura em Python puro.
- O fallback via Excel/PowerShell só é usado localmente se um `.xls` vier fora do padrão esperado.
- No Streamlit Cloud, o ideal é trabalhar com a leitura Python nativa, sem depender do Excel instalado.
- Os arquivos reais de exemplo (`Resumo Mensal*.xls` e `DOMINIO LANCAMENTO*.csv`) estão no `.gitignore` por segurança e privacidade.

## Execução local

```powershell
streamlit run streamlit_app.py
```

Depois abra:

```text
http://localhost:8501
```

## Fluxo

1. Envie o arquivo principal `Resumo Mensal` da filial.
2. Se houver 13º separado, envie também o arquivo complementar `Resumo Mensal13`.
3. A aplicação identifica as competências e calcula as bases da folha.
4. Os lançamentos aparecem na tela com memória de cálculo.
5. O mapeamento pode ser editado sem reler as planilhas.
6. O resultado pode ser exportado em `CSV` ou `XLSX`.

## Validação local

```powershell
python verify_sample.py
python -m py_compile dominio_parser.py streamlit_app.py verify_sample.py
```

Observação:

- `verify_sample.py` depende dos arquivos locais de exemplo e foi pensado para uso na sua máquina, não como parte obrigatória do deploy no Streamlit Cloud.

## GitHub

Esta pasta ainda não estava versionada em Git quando o projeto foi ajustado. Para publicar:

```powershell
git init
git add .
git commit -m "Primeira versão Streamlit do gerador Domínio Lançamento"
```

Depois crie um repositório no GitHub e conecte o remoto:

```powershell
git remote add origin https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git
git branch -M main
git push -u origin main
```

## Streamlit Cloud

Depois que o repositório estiver no GitHub:

1. Acesse `https://share.streamlit.io/` ou o fluxo atual do Streamlit Community Cloud.
2. Escolha o repositório.
3. Defina o arquivo principal como `streamlit_app.py`.
4. Faça o deploy.

## Decisão de regra importante

Na filial 44, a aplicação foi ajustada para permanecer fiel ao `Resumo Mensal 44.xls`, mesmo que isso não reproduza exatamente o `DOMINIO LANCAMENTO - 44.csv` original, porque esse CSV contém linhas que não são inferíveis com segurança a partir do relatório-fonte.
