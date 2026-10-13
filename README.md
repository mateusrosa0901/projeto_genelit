# Genesis v4.0 — Repasse de Parceiros Digitais e Polos

> ℹ️ A **v4.1** ([genesis-v4.1-repasse-parceiro-polo](https://github.com/ItaloHenriqueMarquesCampos/genesis-v4.1-repasse-parceiro-polo)) mantém tudo desta versão
> e acrescenta o repasse do Parceiro Polo que vende em outro polo.

Calcula, a partir do relatório **Receita Polo** do Pincel, o repasse de cada **Parceiro Digital
(Parceiro Comercial)** e de cada **Polo**, seguindo o *Manual de Auditoria – Repasse e Premiação*,
e gera os arquivos no mesmo layout das planilhas feitas à mão.

Feito em Python com bibliotecas simples de análise de dados: **pandas** e **openpyxl**.

## Como rodar (todo mês)

1. No Pincel: **Gestão > Receita Polo > Geral > Mês/Ano > Gerar Relatório**.
2. Salve o arquivo na pasta `entrada` como `relatorioReceitaPolo_convertido.xlsx`.
3. Dê dois cliques em **`EXECUTAR REPASSE.bat`** (Windows) ou **`EXECUTAR REPASSE.command`** (Mac),
   ou rode no terminal, dentro da pasta do projeto:

   ```
   pip install -r requirements.txt     (só na primeira vez)
   python main.py
   ```

Leva menos de 1 minuto. O resultado sai em `saida/<ANO>/<MM>. REF. <MÊS>/`.

## Versão web (Streamlit)

`streamlit_app.py` faz o mesmo cálculo no navegador: envie o `relatorioReceitaPolo_convertido.xlsx` e baixe
a pasta do mês em `.zip` (os mesmos arquivos de `saida/<ANO>/<MM>. REF. <MÊS>/`). Os arquivos enviados ficam
numa pasta temporária só durante o cálculo e são apagados em seguida. O `D_Comissao.csv` vem do repositório.

```
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Os cadastros (`PLANILHA GERAL.xlsx`, `D_Polos.csv`, `D_Excecoes.csv` e, opcional, `D_Parceiros_Ajustes.csv`)
podem vir de dois lugares, conforme o `.streamlit/secrets.toml` (modelo em `.streamlit/secrets.exemplo.toml`;
o arquivo de verdade está no `.gitignore`):

* **Upload na página** — sem as seções `[drive]` e `[gcp_service_account]` no `secrets.toml`.
* **Pasta do Google Drive** — com as duas seções preenchidas:
  1. No [Google Cloud Console](https://console.cloud.google.com/): crie um projeto, ative a **Google Drive API**
     e crie uma **conta de serviço** com uma chave **JSON**.
  2. No Drive, crie uma pasta com os arquivos acima (com esses nomes; deixe a PLANILHA GERAL como `.xlsx`, sem
     converter para Planilhas Google, para manter a formatação) e compartilhe a pasta, como **Leitor**, com o
     `client_email` da conta de serviço.
  3. No `secrets.toml`: `pasta_id` (final do link da pasta) em `[drive]` e os campos do JSON em
     `[gcp_service_account]`. No Streamlit Community Cloud o mesmo conteúdo vai em *Settings > Secrets*.

  Com o Drive ligado, quem abre a página gera arquivos com os dados do cadastro (CPFs, e-mails): preencha
  `senha_app` no `secrets.toml` para a página pedir senha.

## O que sai

| Arquivo / pasta | Para que serve |
|---|---|
| `PLANILHA GERAL - <MÊS>-<ANO>.xlsx` | A planilha de controle de sempre (mesma formatação, fórmulas e colunas), com o **Valor** de cada Parceiro Digital calculado. RPA, POLO, DISTRATO e NÃO ENVIAR REPASSE ficam zerados nela, como no manual. |
| `RELATÓRIOS PARCEIROS DIGITAIS/<nome>.xlsx` | Um relatório por parceiro com repasse, **idêntico ao feito à mão**: aba `Página1`, as 18 colunas, `VALOR DO POLO` oculto, `VALOR DO PARCEIRO` com fórmula (`=K2*0.2`, `=K2*0.1`, `=K2`, `0`), linha **TOTAL**. Os parceiros pagos via RH ficam em `RPA (RH)`. |
| `RELATÓRIO DE PAGAMENTOS EM POLOS ATIVOS/` e `... DESATIVADOS (AGUARDANDO AUTORIZAÇÃO)/` | Um relatório por polo (layout de Açailândia). |
| `TOTAL REPASSE - <MÊS>-<ANO>.xlsx` | Resumo, **Parceiros Digitais**, **RPA (RH)**, Parceiros Não Pagos, Vendedores Não Cadastrados, Cruzamento de Nomes, **Conferência Parceiros** (manual × calculado), Repasse Polos, Pendências Polos, Memória de Cálculo, Regras e Exceções. |
| `F_pagamentos.csv` | Base linha a linha para o Power BI. |

## Regras dos Parceiros Digitais

Conferidas linha a linha com os relatórios manuais de **setembro/2026**: as 2.623 linhas das 7 planilhas
de parceiros batem com o cálculo (mesmas linhas, mesmo percentual, mesmo total).

* **Quem é Parceiro Digital:** quem está na **PLANILHA GERAL** (pasta `dados`) ou em
  `dados/D_Parceiros_Ajustes.csv`. O `NOME PARCEIRO` do Pincel é ligado ao cadastro pelo nome igual, pelo
  CPF/CNPJ, por nome contido ou por nome muito parecido. Todo cruzamento que não é "nome igual" aparece na
  aba *Cruzamento de Nomes*.
* **Responsável de polo também é Parceiro Digital:** quem é responsável de polo em `D_Polos.csv`, vendeu
  cadastrando o aluno em **outro** polo e não está na PLANILHA GERAL entra automaticamente no fim dela
  (coluna Q = `INCLUÍDO PELO PROGRAMA (RESPONSÁVEL DE POLO)`) e ganha relatório de parceiro com essas vendas.
  A venda no próprio polo continua no repasse do polo. A ligação é pelo CPF/CNPJ do responsável ou, sem
  documento em comum, pelo nome — nesse caso só se o vendedor tiver EQUIPE de parceiro no Pincel.
  Liga/desliga e status de polo considerados: `RESPONSAVEL_DE_POLO_E_PARCEIRO` e
  `STATUS_RESPONSAVEL_PARCEIRO` em `config.py`. Para não pagar alguém: `situacao` = `NÃO PAGAR` em
  `D_Parceiros_Ajustes.csv`.
* **Percentual** (`D_Comissao.csv`, papel `PARCEIRO_COMERCIAL`, pela **data de matrícula**):
  * `MAT` (matrícula) = 100%
  * `MAT-R` (rematrícula) = 0%
  * parcela `1` = 0%
  * mensalidades e `NEG` = 10% para matrícula até 31/01/2023 e 20% a partir de 01/02/2023
* **Exceções** (`D_Excecoes.csv`):
  * Danubia Kelly Nunes Bruschi e Integral Representações (Kester): 100% na **1ª mensalidade = PARCELA 2**
    de Graduação/Tecnólogo. Nesse caso o polo não recebe sobre a parcela, porque polo + parceiro não
    passam de 100%.
  * Marcieni Alves Vieira: 40% fora da Graduação para matrículas a partir de 01/2024.
  * Wilhan, Ulisses e Altivo: exceções do repasse dos **polos** deles.
* **Não entram no relatório do parceiro:**
  * vendas no **polo ATIVO do qual ele é o responsável**, porque já estão no repasse do polo (ex.: Wilhan
    em Belo Oriente/Guanhães/Santana do Paraíso, Marcieni em São Francisco (Jardim Graziela));
  * matrículas anteriores à **data de início da parceria** (`D_Parceiros_Ajustes.csv`).
* **Situação na PLANILHA GERAL** (coluna `PF / PJ` e coluna de observação, sem título — coluna Q):
  * `POLO`, `DISTRATO...` e `NÃO ENVIAR REPASSE` não recebem;
  * `RPA` / "Retornou como parceiro" recebe pelo RH (aba *RPA (RH)*, com 20% de INSS patronal);
  * os demais recebem.

## Cadastro de Parceiros Digitais

* **`dados/PLANILHA GERAL.xlsx`** — a planilha de controle de sempre, que é o cadastro. Parceiro novo:
  inclua uma linha. Parceiro que deixou de receber: escreva a situação na coluna Q, por exemplo
  `DISTRATO`, `NÃO ENVIAR REPASSE` ou `INATIVO`. O programa usa essa planilha como modelo da
  `PLANILHA GERAL - <MÊS>-<ANO>.xlsx` e só troca o Valor, sem mudar a formatação.
* **`dados/D_Parceiros_Ajustes.csv`** (abre no Excel), uma linha por parceiro que precisa de ajuste:

| coluna | uso |
|---|---|
| `nome_planilha` | nome como está na PLANILHA GERAL (se não estiver lá, o parceiro é incluído no fim) |
| `apelidos_pincel` | outros nomes do parceiro no Pincel, separados por `\|` |
| `cpf_cnpj_extra` | outros CPF/CNPJ do parceiro, separados por `\|` |
| `data_inicio_parceria` | matrículas antes desta data não entram (dd/mm/aaaa) |
| `situacao` | força a situação: `PAGAR`, `RPA` ou `NÃO PAGAR` |
| `cpf_cnpj`, `pf_pj`, `email` | só para parceiro que ainda não está na PLANILHA GERAL |
| `observacao` | anotação livre |

## Conferência com planilhas feitas à mão (opcional)

Coloque em `referencias/` as planilhas calculadas à mão **do mesmo mês**: relatórios de parceiro,
a PLANILHA GERAL manual e relatórios de polo. Elas **não entram no cálculo**. A aba
*Conferência Parceiros* mostra, parceiro a parceiro, o valor manual × o calculado com a explicação
de cada diferença. A aba *Conferência Linhas Parceiros* mostra a mesma comparação linha a linha.
Ao virar o mês, tire as do mês anterior.

## Arquivos de entrada

| Arquivo | Para que serve | Quando mexer |
|---|---|---|
| `entrada/relatorioReceitaPolo_convertido.xlsx` | Pagamentos do mês (Pincel) | Todo mês |
| `dados/PLANILHA GERAL.xlsx` | Cadastro de Parceiros Digitais | Parceiro novo / distrato |
| `dados/D_Parceiros_Ajustes.csv` | Apelidos, data de início, parceiros fora da planilha | Quando precisar |
| `dados/D_Comissao.csv` | Percentuais por vigência, modalidade e tipo de parcela | Mudança de contrato |
| `dados/D_Excecoes.csv` | Exceções do manual | Nova exceção |
| `dados/D_Polos.csv` | Polos, responsáveis e situação (ATIVO/DESATIVADO/PARCEIRO_SEM_POLO) | Mudança de polo |
| `config.py` | Pastas, nomes de arquivo e opções | Raramente |

## O que mudou em relação ao `genesis.py` (regularização)

1. **Parceiro Digital passou a ser quem está na PLANILHA GERAL.** Antes, todo vendedor não identificado
   como interno virava parceiro (482 parceiros, R$ 995.799,41 em set/2026, com centenas de vendedores
   internos sem EQUIPE). Agora são **225 parceiros** com repasse: 197 na PLANILHA GERAL
   (R$ 340.131,34) e 28 via RPA (R$ 4.813,85).
2. **"1ª mensalidade" das exceções = PARCELA 2** (antes era a parcela 1). Danubia e Integral recebem
   100% da parcela 2 de Graduação/Tecnólogo, e Ulisses idem no polo de Cuiabá.
3. **Parcela 1 e MAT-R sem repasse ao parceiro**, como nos relatórios manuais.
4. **Venda no próprio polo ATIVO** fica no repasse do polo, e não no do parceiro.
5. **Relatório por parceiro idêntico ao manual** e **PLANILHA GERAL do mês** gerada a partir da planilha
   de controle.
6. **Conferência automática** com as planilhas feitas à mão, com a explicação de cada diferença.
7. Código separado em módulos (`genesis/`), com as configurações em `config.py`.

## Estrutura

```
Genesis/
├── main.py                  ← rode este arquivo
├── streamlit_app.py         ← versão web (streamlit run streamlit_app.py)
├── config.py                ← configurações
├── .streamlit/              ← secrets.toml (fora do GitHub) e o modelo secrets.exemplo.toml
├── EXECUTAR REPASSE.bat     ← atalho Windows
├── EXECUTAR REPASSE.command ← atalho Mac
├── requirements.txt
├── genesis/                 ← código
│   ├── entradas.py          leitura do Pincel, D_*.csv e PLANILHA GERAL
│   ├── polos.py             repasse dos polos
│   ├── parceiros.py         repasse dos Parceiros Digitais
│   ├── excel.py             relatórios no layout manual e PLANILHA GERAL
│   ├── consolidado.py       TOTAL REPASSE (resumo, RPA, conferências...)
│   ├── conferencia.py       comparação com as planilhas feitas à mão
│   ├── processo.py          sequência do cálculo e dos arquivos
│   ├── web.py               execução da versão web (pasta temporária e .zip)
│   ├── drive.py             leitura dos cadastros no Google Drive (versão web)
│   └── util.py              funções de apoio
├── dados/                   ← cadastros e regras
├── entrada/                 ← relatório do Pincel do mês
├── referencias/             ← (opcional) planilhas feitas à mão do mês
└── saida/                   ← resultado
```

Requisitos: Python 3.9 ou superior, `pandas` e `openpyxl`.

## Primeira vez a partir do GitHub

O repositório não guarda dados pessoais (relatório do Pincel, PLANILHA GERAL, CPFs, planilhas de
referência). Depois de clonar, copie para a pasta `dados` os seus arquivos reais (`PLANILHA GERAL.xlsx`,
`D_Polos.csv`, `D_Excecoes.csv`, `D_Parceiros_Ajustes.csv`) — os modelos em `dados/modelos/` mostram o
formato de cada um.

## Família Genesis

| Versão | Repositório | Situação |
|---|---|---|
| v1 | [genesis-v1-comissao-parceiros](https://github.com/ItaloHenriqueMarquesCampos/genesis-v1-comissao-parceiros) | arquivada |
| v2 | [genesis-v2-repasse-polos](https://github.com/ItaloHenriqueMarquesCampos/genesis-v2-repasse-polos) | arquivada |
| **v4.0** | genesis-v4.0-repasse-parceiros-digitais (este) | ativa |
| v4.1 | [genesis-v4.1-repasse-parceiro-polo](https://github.com/ItaloHenriqueMarquesCampos/genesis-v4.1-repasse-parceiro-polo) | **atual** |
