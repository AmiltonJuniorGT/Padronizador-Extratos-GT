# Padronizador de Extratos GT

Aplicação estática para converter extratos brutos do Itaú em duas abas:

- RAW_EXTRATO
- PIVOT_MENSAL

## Como usar

1. Acesse a página publicada no GitHub Pages.
2. Selecione ou arraste o arquivo `.xlsx`, `.xls` ou `.csv`.
3. Clique em `Processar arquivo`.
4. Confira os previews.
5. Clique em `Exportar XLSX padronizado`.

## Formato gerado

### RAW_EXTRATO

| Data | Denominação | Valor | Mes |

### PIVOT_MENSAL

| Denominação | 2023-02 | 2023-03 | ... |

## Observações

- O app detecta automaticamente a linha de cabeçalho.
- O extrato precisa conter campos parecidos com Data, Lançamento/Descrição/Histórico e Valor.
- Valores negativos são tratados como positivos para análise de despesas.
