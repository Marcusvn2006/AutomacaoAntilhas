# Pipeline de Automação — Antilhas

Automatiza o preenchimento semanal das **Planilhas Pai** de controle de estoque de embalagens
para as regiões **Jaú, Bauru e Praia Grande** — 32 lojas Boticário no total.

O pipeline é composto por dois processos independentes que rodam em sequência toda semana.

---

## Como funciona

```
Processo 1 (rodar.bat)
  └── Abre as 3 Planilhas Pai e insere as colunas USOU da semana
  └── Preenche valores de referência (AUXILIAR.xlsx / AUXILIAR_VDS.xlsx)
  └── Avança referências, esconde semanas antigas, gera backup

Processo 2 (rodar_p2.bat)
  └── Lê os arquivos .xlsm/.xlsx enviados pelas gerentes (pasta PENDENTES)
  └── Aplica regras de conversão PCT (unidades → caixas)
  └── Grava os valores nas linhas corretas das Planilhas Pai
  └── Move arquivos processados para PROCESSADOS/
```

---

## Pré-requisitos

- Python 3.10 ou superior
- Pacotes: `openpyxl` e `pyyaml`

```
pip install openpyxl pyyaml
```

> ⚠️ Na instalação do Python, marque **"Add Python to PATH"**

---

## Estrutura de pastas

```
antilhas/
├── rodar.bat                    ← executa o Processo 1 (duplo clique)
├── rodar_p2.bat                 ← executa o Processo 2 (duplo clique)
├── processo1.py
├── processo2.py
├── config_p1.yaml               ← configuração do Processo 1 (não sobe pro Git)
├── config_p2.yaml               ← configuração do Processo 2 (não sobe pro Git)
├── config_p1.yaml.example       ← modelo do config_p1.yaml
├── config_p2.yaml.example       ← modelo do config_p2.yaml
├── escritor.py
├── validadores.py
├── processadores/
│
├── 00_ENTRADA/
│   ├── AUXILIAR.xlsx            ← fonte de referência (não sobe pro Git)
│   ├── AUXILIAR_VDS.xlsx        ← fonte de referência (não sobe pro Git)
│   └── AUXILIAR_EMAIL/
│       ├── BAURU/
│       │   ├── PENDENTES/       ← arquivos das gerentes entram aqui
│       │   ├── PROCESSADOS/     ← movidos após processamento bem-sucedido
│       │   └── ERROS/           ← movidos em caso de erro
│       ├── JAU/
│       │   ├── PENDENTES/
│       │   ├── PROCESSADOS/
│       │   └── ERROS/
│       └── PRAIA/
│           ├── PENDENTES/
│           ├── PROCESSADOS/
│           └── ERROS/
│
├── 01_BACKUP/                   ← backup automático das Planilhas Pai
├── 02_LOGS/                     ← logs de execução
└── 04_PLANILHAS_PAI/            ← Planilhas Pai .xlsm (não sobem pro Git)
```

---

## Configuração inicial

Copie os arquivos de exemplo e preencha com os dados reais:

```
config_p1.yaml.example  →  config_p1.yaml
config_p2.yaml.example  →  config_p2.yaml
```

> Os arquivos `.yaml` reais contêm dados internos da empresa e estão no `.gitignore`.
> Os arquivos `.example` são modelos com dados fictícios para referência.

---

## Execução semanal

### Etapa 1 — Processo 1 (início da semana)

1. Feche o Excel
2. Duplo clique em `rodar.bat`
3. Aguarde **"Concluido com sucesso!"**

### Etapa 2 — Receber arquivos das gerentes

Baixe os arquivos enviados por e-mail, renomeie conforme a tabela abaixo e coloque na pasta `PENDENTES/` da região correspondente:

| Região | Exemplos de arquivo |
|:---:|:---|
| BAURU | `BSH.xlsm`, `BOUL.xlsm`, `TT.xlsm`, `CDB.xlsm`, `ERB.xlsm` ... |
| JAU | `JC.xlsm`, `BB.xlsm`, `ERJ.xlsm`, `CDJ.xlsm`, `ER SM.xlsm` ... |
| PRAIA | `PL.xlsm`, `BOQ.xlsm`, `CDP.xlsm`, `ER BOQ.xlsm`, `ER MG.xlsm` ... |

> O nome do arquivo (sem extensão) deve bater exatamente com a chave no `config_p2.yaml`.
> Maiúscula/minúscula não importa. Espaços no meio **importam** (ex: `ER SM.xlsm`).

### Etapa 3 — Processo 2 (após receber as planilhas)

1. Feche o Excel
2. Confirme que os arquivos estão nas pastas `PENDENTES/`
3. Duplo clique em `rodar_p2.bat`
4. Aguarde **"Concluido com sucesso!"**

---

## Verificar resultados

| O que checar | Onde |
|:---|:---|
| Lojas processadas | Mensagens na janela preta |
| Arquivos com sucesso | Pasta `PROCESSADOS/` da região |
| Arquivos com erro | Pasta `ERROS/` — arquivo `_ERRO.txt` explica o problema |
| Lojas pendentes | Ainda em `PENDENTES/` (não foram movidas) |
| Log detalhado | Pasta `02_LOGS/` — arquivo `.log` com data de hoje |

---

## Problemas comuns

| Problema | Causa | Solução |
|:---|:---|:---|
| "Python não é reconhecido" | Python não está no PATH | Reinstale marcando "Add to PATH" |
| "No module named openpyxl" | Biblioteca não instalada | `pip install openpyxl pyyaml` |
| Arquivo ficou em PENDENTES | Nome do arquivo errado | Renomeie conforme a tabela e rode de novo |
| "Arquivo aberto no Excel" | Excel estava aberto | Feche o Excel e rode de novo |
| Processo 2 bloqueado | Execução anterior travou | Delete `.processo2.lock` na pasta `antilhas/` |
| "Processo 1 não rodou esta semana" | P1 não foi executado | Rode `rodar.bat` primeiro |

---

## Adicionar nova loja

1. Abra `config_p2.yaml`
2. Adicione a loja na seção correspondente à região:

```yaml
SIGLA: { aba_pai: "SIGLA 00000", regiao: regiao, <<: *PADRAO }
```

3. Para lojas CD/ER com mapeamento próprio, siga o padrão dos blocos `ERJ` ou `CDJ` já existentes
4. O arquivo da gerente deve ser nomeado exatamente com a chave usada (ex: `SIGLA.xlsm`)

---

## Segurança

Os arquivos sensíveis estão protegidos pelo `.gitignore` e **nunca sobem para o repositório**:

- `config_p1.yaml` e `config_p2.yaml` — contêm dados internos da empresa
- `*.xlsm` / `*.xlsx` — Planilhas Pai e arquivos das lojas
- `01_BACKUP/`, `02_LOGS/`, `00_ENTRADA/` — dados operacionais
- `.processo2.lock`, `.env`
