# 💊 PMI – Sistema de Controle de Estoque de Remédios em VBA

Projeto final da disciplina de PMI, desenvolvido em **VBA (Visual Basic for Applications)** com interface no Excel. O sistema permite **cadastrar**, **consultar**, **monitorar**, **vender**, **remover** e **acompanhar movimentações** de medicamentos, com foco em controle de validade, estoque e regras de comercialização (tarja).

## 🎯 Objetivo

Desenvolver uma solução automatizada em VBA para gerenciamento de estoque de medicamentos, com validações de entrada, controle de vencimento e movimentações seguras, seguindo boas práticas de manipulação de dados no Excel.

## 🧭 Interface do Sistema

A planilha conta com uma interface intuitiva, contendo **botões visuais**, **instruções claras** e **validações automáticas** para facilitar a experiência do usuário.

### 🔘 Botões disponíveis

- **Cadastro** – Realiza o cadastro de novos produtos com validação completa.
- **Consulta** – Permite buscar produtos por **código** ou **nome**.
- **Monitoramento** – Exibe medicamentos com problemas de **validade** ou **estoque**.
- **Venda/Remoção** – Realiza a saída de produtos com registro automático.
- **Movimentação** – Mostra todas as ações de entrada e saída registradas.

## 📝 Instruções para Cadastro

- Preencha os dados na **linha 2** da aba *Cadastro*:
  - Nome do produto
  - Tipo de tarja (vermelha, preta ou livre)
  - Quantidade
  - Data de validade
  - Fornecedor
  - Localização (Prateleira/Setor)
- Campos obrigatórios e formatados incorretamente são **rejeitados automaticamente**.
- O **código do produto** é gerado automaticamente após o cadastro.
- A **data de validade** deve ser superior à data atual.

### 🟦 Tipos de Tarja

- 🔴 **Vermelha**: Controle com receita médica
- ⚫ **Preta**: Uso restrito, prescrição especial
- ⚪ **Livre**: Venda sem receita

## 🔍 Consulta

- A consulta pode ser feita por **código** ou **nome do produto**.
- Produtos duplicados por nome exibem todas as versões com suas respectivas validades e localizações.

## 🚨 Verificação de Estoque

O sistema realiza uma verificação automática na planilha de estoque e colore os itens de acordo com o estado crítico:

| Cor        | Situação Identificada                     |
|------------|-------------------------------------------|
| 🔴 Vermelho | Estoque baixo **e** produto próximo do vencimento |
| 🟡 Amarelo  | Produto perto da validade                |
| 🟠 Laranja  | Estoque abaixo de 20%                    |
| ⚫ Preto    | Produto vencido (bloqueado para venda)   |

## 📦 Venda e Remoção

- A **venda** exige a inserção do código e quantidade.
- Produtos vencidos **não podem ser vendidos**.
- A **remoção** elimina completamente um produto, registrando os dados na aba *Movimentação*.

## 📈 Monitoramento

- A aba **Monitoramento** é alimentada automaticamente com os produtos que:
  - Estão com validade próxima ou vencidos
  - Têm estoque abaixo de 20%
- Os produtos são exibidos com **cores de alerta** e **detalhes completos**.

## 🗂️ Estrutura do Projeto

```
PMI/
├── Projeto PMI.xlsm           # Planilha com as macros e botões
├── Cadastro.bas               # Código de cadastro de produtos
├── Consulta.bas               # Código de consulta por nome/código
├── Monitoramento.bas          # Código que filtra e colore produtos críticos
├── Retirada.bas               # Código para venda e remoção
└── README.md                  # Este arquivo
```

## 💻 Tecnologias Utilizadas

- **VBA (Visual Basic for Applications)**
- **Excel (planilhas, botões, validação de dados)**
- Interface 100% funcional dentro do Excel

## 👨‍🎓 Aluno Responsável

- **Guilherme Machado da Silva**
