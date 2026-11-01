# Testes e validação do Relatório Fácil

## 1. Objetivo

Este documento registra os testes realizados durante o desenvolvimento do Relatório Fácil. Os testes foram feitos com dados fictícios e cenários controlados. O status indicado aqui corresponde ao que foi observado até a última rodada de validação documentada e deverá ser atualizado quando novas verificações forem concluídas.

## 2. Testes funcionais

| Teste | Comportamento esperado | Resultado observado | Status |
|---|---|---|---|
| Cadastro de produto | Registrar um produto com os campos obrigatórios válidos | Produto de teste cadastrado pela interface do Excel | Aprovado |
| SKU duplicado | Impedir o cadastro de um SKU já existente | A macro exibiu aviso de SKU já cadastrado | Aprovado |
| Registro de venda | Registrar venda com SKU, quantidade, data e pagamento válidos | Venda de teste registrada na tabela `tbVendas` | Aprovado |
| SKU inexistente | Recusar venda para produto não cadastrado | A macro identificou SKU não encontrado | Aprovado |
| Estoque insuficiente | Impedir quantidade superior ao saldo estimado | A macro exibiu aviso de estoque insuficiente | Aprovado |
| Desconto inválido | Recusar desconto negativo ou superior ao subtotal | Cenário de desconto inválido foi testado e recusado pelo sistema | Aprovado |
| Atualização do Top 10 | Consolidar produtos vendidos no período selecionado | Ranking foi atualizado e passou a apresentar os produtos do período de teste | Aprovado |
| Geração de PDF | Gerar arquivo PDF a partir do Excel e do programa Python | PDF de duas páginas gerado e aberto com dados de teste | Aprovado |

## 3. Validação dos cálculos financeiros

Foi utilizado um cenário controlado com três vendas:

| Produto | Quantidade | Preço unitário | Desconto | Pagamento | Total da venda | Custo total |
|---|---:|---:|---:|---|---:|---:|
| Produto A | 2 | R$ 20,00 | R$ 0,00 | Pix | R$ 40,00 | R$ 20,00 |
| Produto B | 1 | R$ 50,00 | R$ 10,00 | Cartão de crédito | R$ 40,00 | R$ 30,00 |
| Produto C | 3 | R$ 12,00 | R$ 0,00 | Dinheiro | R$ 36,00 | R$ 15,00 |
| **Total** | **6** |  | **R$ 10,00** |  | **R$ 116,00** | **R$ 65,00** |

Indicadores esperados e observados na planilha:

| Indicador | Valor esperado | Valor observado | Status |
|---|---:|---:|---|
| Faturamento | R$ 116,00 | R$ 116,00 | Aprovado |
| Custo dos produtos vendidos | R$ 65,00 | R$ 65,00 | Aprovado |
| Lucro bruto estimado | R$ 51,00 | R$ 51,00 | Aprovado |
| Margem bruta estimada | 43,97% | 43,97% | Aprovado |
| Unidades vendidas | 6 | 6 | Aprovado |
| Registros de venda | 3 | 3 | Aprovado |
| Ticket médio | R$ 38,67 | R$ 38,67 | Aprovado |

### Conferência por forma de pagamento

| Forma de pagamento | Total esperado | Total observado | Status |
|---|---:|---:|---|
| Pix | R$ 40,00 | R$ 40,00 | Aprovado |
| Cartão de crédito | R$ 40,00 | R$ 40,00 | Aprovado |
| Dinheiro | R$ 36,00 | R$ 36,00 | Aprovado |
| Cartão de débito | R$ 0,00 | R$ 0,00 | Aprovado |
| Outros | R$ 0,00 | R$ 0,00 | Aprovado |

## 4. Teste com base demonstrativa maior

Também foi gerada uma base demonstrativa fictícia abrangendo setembro e outubro de 2026. O relatório consolidado apresentado durante essa validação informou:

| Indicador | Resultado apresentado |
|---|---:|
| Registros de vendas | 180 |
| Unidades vendidas | 329 |
| Faturamento | R$ 15.812,58 |
| Custo dos produtos vendidos | R$ 7.255,90 |
| Lucro bruto estimado | R$ 8.556,68 |
| Ticket médio | R$ 87,85 |
| Margem bruta estimada | 54,1% |

O relatório também apresentou ranking de produtos e identificação de três produtos abaixo do estoque mínimo. Esses valores foram exibidos no PDF demonstrativo e são derivados de dados fictícios.

A divisão exata das 180 vendas entre setembro e outubro não está registrada neste documento como uma validação independente por mês.

## 5. Teste de período sem vendas

Foi selecionado um período sem registros de vendas (02/11/2026). A rotina de atualização do Top 10 exibiu a mensagem:

> Não existem vendas no período selecionado.

**Status: validação parcial.** A mensagem confirma que a rotina identificou a ausência de vendas. Ainda é necessário confirmar se, após essa mensagem, o processo completo gera o PDF com indicadores zerados e ranking vazio.

## 6. Pontos que permanecem em validação

- Confirmar a geração completa do PDF para um período sem vendas, com indicadores zerados e ranking vazio.
- Testar separadamente os períodos de setembro e outubro de 2026 e conferir a soma com o consolidado.
- Continuar os testes de apresentação visual dos gráficos e tabelas.
- Validar a versão de demonstração em uma cópia limpa, separada dos dados usados no desenvolvimento.

## 7. Observações sobre os resultados

Os testes foram realizados em ambiente de desenvolvimento com Microsoft Excel desktop no Windows, macros VBA e programa auxiliar Python. Os resultados refletem os cenários e arquivos utilizados na ocasião; não representam uma certificação independente nem garantem compatibilidade com todas as versões do Excel ou configurações de segurança do Windows.

O lucro bruto é estimado e não considera automaticamente despesas operacionais, impostos, taxas, devoluções, perdas ou outros custos. O estoque é calculado na versão atual como estoque inicial menos vendas, sem controle completo de reposições e ajustes.

Todos os dados apresentados neste documento são fictícios e utilizados exclusivamente para testes e demonstração.
