# Painel UPAs · Antes x Pós-Projeto

Painel em Streamlit que lê a planilha **ANALISE_UPAS.xlsx** (uma aba por mês) e mostra, para cada uma das 13 UPAs, o que melhorou e o que piorou depois do início do projeto.

O painel se atualiza sozinho: ele relê a planilha a cada 5 minutos. Também dá para forçar a leitura pelo botão **Atualizar dados agora**.

## O que o painel mostra

| Aba | Conteúdo |
|---|---|
| Visão geral | Cartões da rede, mapa de calor (UPA × indicador) e saldo de indicadores que melhoraram ou pioraram por UPA |
| Por UPA | Tabela Antes x Pós da UPA escolhida e evolução mensal, com a média da rede e a data de início do projeto |
| Por indicador | Comparação Antes x Pós de todas as UPAs para um indicador |
| Tabela completa | Todos os números, com download em CSV |

## Como cada indicador é avaliado

O que conta como melhora segue as metas das UPAs no **51º Termo Aditivo ao Contrato nº 001/2018-SES/DF**:

| Indicador | Melhora quando |
|---|---|
| Atend. médicos + urgência com observação ≤ 24h (indicador da meta) | aumenta |
| Atend. Médico em UPA, Atend. Urgência Obs. ≤ 24h, total de atendimentos | aumenta |
| % de GAEs classificadas | aumenta |
| Tempo médio Laranja, Amarelo e Verde | diminui |
| Acolhimento, exames, RX, ECG, GAEs abertas e classificadas | sem meta no contrato: aparece só a variação. Dá para mudar isso na barra lateral |

- **Antes / Pós:** o painel usa a coluna **Fase** da planilha (Antes-Projeto / Pós-Projeto).
- **Volumes:** média por dia do período.
- **Tempos:** média do período. Tempos zerados são tratados como sem informação.
- **% de GAEs:** soma das classificadas ÷ soma das abertas no período.
- **Estável:** variação menor que o limite escolhido na barra lateral (padrão: 5%).

## Publicar no Streamlit Community Cloud

1. Crie um repositório no GitHub e envie estes arquivos: `app.py`, `requirements.txt`, `.streamlit/config.toml` e, se quiser, a pasta `dados/`.
2. Em [share.streamlit.io](https://share.streamlit.io), clique em **Create app**, escolha o repositório e indique o arquivo `app.py`.
3. Em **Advanced settings → Secrets** (ou depois, em **Settings → Secrets**), cole o link da planilha:
   ```toml
   PLANILHA_URL = "https://docs.google.com/spreadsheets/d/SEU_ID/edit"
   ```
4. Clique em **Deploy**.

### Links aceitos

A planilha precisa estar compartilhada como **"Qualquer pessoa com o link pode ver"**.

- **Google Sheets:** cole o link normal de edição. O painel converte para o download em Excel.
- **OneDrive / SharePoint:** use o link de compartilhamento. O painel acrescenta `download=1`. Em SharePoint corporativo, o acesso anônimo pode estar bloqueado pela organização.
- **Dropbox** ou **GitHub** (link do arquivo no repositório).

Sem link configurado, o painel usa o arquivo `dados/ANALISE_UPAS.xlsx` do repositório. Nesse caso, ele se atualiza toda vez que você substituir esse arquivo no GitHub. Também dá para enviar um .xlsx manualmente pela barra lateral.

## Rodar no computador

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Formato esperado da planilha

- Uma aba por mês.
- Linha de cabeçalho com **Data | Fase | Indicador | UPA …**.
- Linha acima do cabeçalho com "PORTE I" / "PORTE III" sobre os grupos de UPAs.
- Os nomes dos indicadores podem variar um pouco: o painel reconhece por palavras-chave (por exemplo, "Radiografias", "Amarelo", "GAEs Abertas").
- Linhas com fórmulas (totais e %) são recalculadas pelo painel a partir dos valores digitados.
- Dias ainda sem nenhum valor são ignorados.
