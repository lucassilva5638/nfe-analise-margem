# Da nota fiscal à margem: DRE e curva ABC automáticas a partir de XMLs de NF-e

Script em Python que lê um lote de XMLs de NF-e e monta sozinho um relatório gerencial em Excel: **DRE mensal, margem por produto, curva ABC e checagem de qualidade dos dados**.

> Todos os dados são **fictícios**. A empresa (Atlas Distribuidora de Suprimentos Industriais Ltda), os CNPJs, os clientes e as notas foram gerados pelo próprio projeto.

## O problema

Em muitas empresas, a margem por produto só aparece no fechamento, quando aparece. Mas os dados já estão nas notas fiscais: preço, quantidade, desconto, ICMS, PIS e COFINS de cada item. Este projeto transforma os XMLs em informação gerencial sem digitar nada.

## O que o script faz

1. **Lê e valida os XMLs** (namespace oficial da NF-e 4.00)
   - remove notas duplicadas
   - ignora notas não autorizadas (cStat ≠ 100)
   - valida o dígito verificador da chave de acesso (módulo 11)
   - confere se a soma dos itens bate com o total da nota
2. **Classifica cada nota** pelo CNPJ da empresa: venda, devolução de venda (finNFe = 4) ou compra
3. **Calcula o custo médio ponderado** de cada produto a partir das compras, líquido do ICMS recuperável
4. **Gera o relatório em Excel**, com fórmulas vivas em cima de uma aba de base de dados:
   - **Resumo:** indicadores e gráficos
   - **DRE Mensal:** receita bruta, devoluções, ICMS, PIS, COFINS, receita líquida, CMV e lucro bruto por mês
   - **Margem por Produto:** margem bruta, preço médio e curva ABC, com destaque para margens abaixo de 15%
   - **Qualidade dos Dados:** tudo o que foi encontrado e como foi tratado
   - **Base Itens:** cada item de nota, pronto para filtro e tabela dinâmica

## Como rodar

```bash
pip install -r requirements.txt
python gerar_xmls_ficticios.py xmls        # cria 515 XMLs fictícios de 2025
python analisar_nfe.py --pasta xmls --cnpj 11222333000181 --saida relatorio_margem_2025.xlsx
```

Para usar com outra empresa, basta apontar `--pasta` para os XMLs e informar o `--cnpj` dela.

## Resultados com os dados fictícios

- Receita líquida de R$ 498 mil e margem bruta de 29,6% no ano
- **Disco flap** e **desengraxante industrial** estão entre os produtos classe A em receita, mas têm margem bruta de 4,5% e 2,4%: vendem muito e quase não deixam resultado
- O analisador encontrou as 3 falhas plantadas no lote: uma nota duplicada, uma nota denegada e uma chave com dígito verificador inválido

## Premissas e simplificações

- Regime: Lucro Presumido (PIS 0,65% e COFINS 3% cumulativos, sem crédito)
- Base de PIS/COFINS sem o ICMS (Tema 69 do STF)
- Custo médio ponderado anual; não trata estoque inicial/final, frete, ICMS-ST nem IPI
- Layout de NF-e simplificado: só os campos usados na análise

## Tecnologias

Python · pandas · openpyxl · xml.etree

## Autor

Lucas — analista fiscal com foco em finanças e FP&A. Projeto de portfólio.
